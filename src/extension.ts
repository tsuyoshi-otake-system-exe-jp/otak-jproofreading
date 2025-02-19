import * as vscode from 'vscode';
import OpenAI from 'openai';
import { HttpsProxyAgent } from 'https-proxy-agent';
import * as https from 'https';

// 基本的な日本語の校正ルール
interface ProofreadingRule {
  pattern: RegExp;
  replacement: string | ((match: string, ...args: string[]) => string);
  description: string;
}

const PROOFREADING_RULES = [
  {
    pattern: /[！？]([^！？])/g,
    replacement: (match: string, p1: string) => match.charAt(0) + '　' + p1,
    description: '感嘆符・疑問符の後には全角スペースを入れる'
  },
  {
    pattern: /([。、．，])\s*([^\n])/g,
    replacement: (match: string, p1: string, p2: string) => p1 + p2,
    description: '句読点の処理'
  },
  {
    pattern: /です([。！？\n]|$)/g,
    replacement: (match: string, p1: string) => 'です。' + (p1 === '\n' ? '\n' : p1 || ''),
    description: '文末の「です」の後には句点が必要'
  },
  {
    pattern: /ます([。！？\n]|$)/g,
    replacement: (match: string, p1: string) => 'ます。' + (p1 === '\n' ? '\n' : p1 || ''),
    description: '文末の「ます」の後には句点が必要'
  }
] as ProofreadingRule[];

// 文体検出
function detectWritingStyle(text: string): { 
  style: 'ですます' | 'である' | 'mixed' | 'unknown';
  details: {
    dearu: number;
    desuMasu: number;
  };
} {
  const dearu = (text.match(/である。|だ。/g) || []).length;
  const desuMasu = (text.match(/です。|ます。/g) || []).length;
  
  let style: 'ですます' | 'である' | 'mixed' | 'unknown';
  if (dearu > 0 && desuMasu > 0) {
    style = 'mixed';
  } else if (dearu > 0) {
    style = 'である';
  } else if (desuMasu > 0) {
    style = 'ですます';
  } else {
    style = 'unknown';
  }

  return { style, details: { dearu, desuMasu } };
}

// 文体に関する修正提案を生成
function getStyleSuggestion(analysis: ReturnType<typeof detectWritingStyle>): string {
  const { style, details } = analysis;
  switch (style) {
    case 'mixed':
      return `文体が混在しています（「である」体: ${details.dearu}箇所、「です・ます」体: ${details.desuMasu}箇所）。一貫した文体の使用を推奨します。`;
    case 'ですます':
      return `現在「です・ます」体で統一されています（${details.desuMasu}箇所）。`;
    case 'である':
      return `現在「である」体で統一されています（${details.dearu}箇所）。`;
    default:
      return '明確な文体が検出されませんでした。';
  }
}

let statusBarItem: vscode.StatusBarItem;
let openai: OpenAI | undefined;
let isProofreading: boolean = false;
let currentProofreadingController: AbortController | undefined;
let currentCancellationTokenSource: vscode.CancellationTokenSource | undefined;

// プロキシ設定を取得する関数
function getProxySettings(): { httpsProxy: string | undefined } {
  const config = vscode.workspace.getConfiguration('otak-jproofreading');
  const configuredProxy = config.get<string>('proxy');

  // 設定されたプロキシがある場合はそれを使用
  if (configuredProxy) {
    return { httpsProxy: configuredProxy };
  }

  // VS Codeのプロキシ設定を使用
  const httpSettings = vscode.workspace.getConfiguration('http');
  const proxy = httpSettings.get<string>('proxy');

  if (proxy) {
    return { httpsProxy: proxy };
  }

  // 環境変数からプロキシ設定を取得
  return {
    httpsProxy: process.env.HTTPS_PROXY || process.env.HTTP_PROXY
  };
}

// OpenAIクライアントを初期化する関数
function initializeOpenAIClient(): OpenAI | undefined {
  const config = vscode.workspace.getConfiguration('otak-jproofreading');
  const apiKey = config.get<string>('openaiApiKey');

  if (!apiKey) {
    return undefined;
  }

  const { httpsProxy } = getProxySettings();
  const configuration: { apiKey: string; httpAgent?: https.Agent; httpsAgent?: https.Agent } = {
    apiKey: apiKey,
  };

  // プロキシが設定されている場合、httpsAgentを設定
  if (httpsProxy) {
    const agent = new HttpsProxyAgent(httpsProxy);
    configuration.httpAgent = agent;
    configuration.httpsAgent = agent;
  }

  return new OpenAI(configuration);
}

// APIキーの設定を促す関数
async function promptForApiKey(): Promise<boolean> {
  const action = await vscode.window.showInformationMessage(
    'OpenAI APIキーが必要です。APIキーを設定しますか？',
    '設定する',
    'キャンセル'
  );

  if (action === '設定する') {
    const apiKey = await vscode.window.showInputBox({
      prompt: 'OpenAI APIキーを入力してください',
      password: true,
      placeHolder: 'sk-...'
    });

    if (apiKey) {
      const config = vscode.workspace.getConfiguration('otak-jproofreading');
      await config.update('openaiApiKey', apiKey, true);
      openai = initializeOpenAIClient();
      vscode.window.showInformationMessage('APIキーを設定しました。');
      return true;
    }
  }
  return false;
}

// ステータスバーの表示を更新
function updateStatusBar(isChecking: boolean = false) {
  if (isChecking) {
    statusBarItem.text = "$(sync~spin) 校正中...";
    statusBarItem.tooltip = "クリックして校正を中止";
    statusBarItem.command = 'otak-jproofreading.cancelProofreading';
    isProofreading = true;
  } else {
    statusBarItem.text = "$(pencil) 校正";
    statusBarItem.tooltip = "クリックして文書を校正";
    statusBarItem.command = 'otak-jproofreading.checkDocument';
    isProofreading = false;
    currentProofreadingController = undefined;
    currentCancellationTokenSource?.dispose();
    currentCancellationTokenSource = undefined;
  }
}

// 校正をキャンセル
function cancelProofreading() {
  currentProofreadingController?.abort();
  currentCancellationTokenSource?.cancel();
  updateStatusBar(false);
}

export function activate(context: vscode.ExtensionContext) {
  // ステータスバーアイテムの作成
  statusBarItem = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Right,
    100
  );
  updateStatusBar();
  statusBarItem.show();

  // OpenAI APIクライアントの初期化
  openai = initializeOpenAIClient();

  // 設定変更を監視
  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration(e => {
      if (e.affectsConfiguration('otak-jproofreading') || e.affectsConfiguration('http')) {
        openai = initializeOpenAIClient();
      }
    })
  );

  // 校正キャンセルコマンド
  let cancelProofreadingDisposable = vscode.commands.registerCommand(
    'otak-jproofreading.cancelProofreading',
    cancelProofreading
  );

  // 文書全体を校正するコマンド
  let checkDocumentDisposable = vscode.commands.registerCommand(
    'otak-jproofreading.checkDocument',
    async () => {
      if (isProofreading) { return cancelProofreading(); }
      const editor = vscode.window.activeTextEditor;
    if (!editor) {
      vscode.window.showWarningMessage('テキストエディタがアクティブではありません。');
      return;
    }

    const document = editor.document;
    const text = document.getText();
    await proofreadText(text, editor);
  });

  // 選択範囲を校正するコマンド
  let checkSelectionDisposable = vscode.commands.registerCommand(
    'otak-jproofreading.checkSelection',
    async () => {
      if (isProofreading) { return cancelProofreading(); }
      const editor = vscode.window.activeTextEditor;
    if (!editor) {
      vscode.window.showWarningMessage('テキストエディタがアクティブではありません。');
      return;
    }

    const selection = editor.selection;
    if (selection.isEmpty) {
      vscode.window.showWarningMessage('テキストが選択されていません。');
      return;
    }

    const text = editor.document.getText(selection);
    await proofreadText(text, editor, selection);
  });

  context.subscriptions.push(statusBarItem);
  context.subscriptions.push(checkDocumentDisposable);
  context.subscriptions.push(cancelProofreadingDisposable);
  context.subscriptions.push(checkSelectionDisposable);
}

// 校正処理のメイン関数
async function proofreadText(
  text: string,
  editor: vscode.TextEditor,
  selection?: vscode.Selection
): Promise<void> {
  try {
    currentProofreadingController = new AbortController();
    currentCancellationTokenSource = new vscode.CancellationTokenSource();
    updateStatusBar(true);

    // プログレス表示
    await vscode.window.withProgress({
      location: vscode.ProgressLocation.Notification,
      title: "校正を開始します",
      cancellable: true
    }, async (progress, token) => {
      token.onCancellationRequested(() => {
        cancelProofreading();
      });

      // 文書全体を取得
      const fullText = editor.document.getText();
      const corrections: { original: string; corrected: string; description: string }[] = [];
      
      // 文脈を考慮したテキストを準備
      let contextBefore = '';
      let contextAfter = '';
      let targetText = text;
      
      if (selection) {
        contextBefore = editor.document.getText(new vscode.Range(
          editor.document.positionAt(0),
          selection.start
        ));
        contextAfter = editor.document.getText(new vscode.Range(
          selection.end,
          editor.document.positionAt(fullText.length)
        ));
      }

      interface RuleBasedResult {
        text: string;
        corrections: Array<{ rule: ProofreadingRule; original: string }>;
        style: ReturnType<typeof detectWritingStyle>;
      }

      interface AIResult {
        text: string;
        response: string | null;
        correction: {
          corrected: string;
          reason: string;
        } | null;
      }

      // ルールベースの校正とAI校正を並行して実行
      const [ruleBasedResult, aiResult]: [RuleBasedResult, AIResult] = await Promise.all([
        // ルールベースの校正
        (async () => {
          progress.report({ message: 'ルールベースの校正を実行中...' });
          let currentText = targetText;
          const ruleCorrections: { rule: ProofreadingRule; original: string }[] = [];

          PROOFREADING_RULES.forEach(rule => {
            const originalText = currentText;
            currentText = currentText.replace(rule.pattern, (match: string, ...args: any[]) => {
              if (typeof rule.replacement === 'function') {
                return rule.replacement(match, ...args);
              }
              return rule.replacement;
            });
            if (originalText !== currentText) {
              ruleCorrections.push({ rule, original: originalText });
            }
          });

          // 文体の一貫性をチェック
          const documentStyle = detectWritingStyle(fullText);
          
          return {
            text: currentText,
            corrections: ruleCorrections,
            style: documentStyle
          };
        })() as Promise<RuleBasedResult>,

        // AI校正
        (async () => {
          if (!openai) {
            const apiKeySet = await promptForApiKey();
            if (!apiKeySet) {
              return { text: targetText, response: null, correction: null };
            }
          }

          if (openai && currentProofreadingController) {
            try {
              progress.report({ message: 'AI校正を実行中...' });
              const aiResponse: OpenAI.Chat.Completions.ChatCompletion = await openai.chat.completions.create({
                model: vscode.workspace.getConfiguration('otak-jproofreading').get('model', 'gpt-4-turbo-preview'),
                messages: [
                  {
                    role: 'system',
                    content: `${selection ? '選択範囲のみを校正し、前後の文脈を考慮して、' : ''}` +
                      `日本語の文章を校正し、以下のJSON形式で返してください：\n` +
                      `文体は既存の文体を維持してください。\n` +
                      '{\n' +
                      '  "corrected": "修正後の文章",\n' +
                      '  "reason": "修正理由の説明"\n' +
                      '}\n'
                  },
                  {
                    role: 'user',
                    content: selection ? 
                      `${contextBefore}【ここから校正対象】${targetText}【ここまで校正対象】${contextAfter}` :
                      targetText
                  }
                ]
              }, {
                signal: currentProofreadingController.signal
              });

              return { text: targetText, response: aiResponse.choices[0].message.content?.trim() || '{}', correction: null };
            } catch (error) {
              console.error('AI校正エラー:', error);
              return { text: targetText, response: null, correction: null };
            }
          }
          return { text: targetText, response: null, correction: null };
        })() as Promise<AIResult>
      ]);

      // 結果を統合
      targetText = ruleBasedResult.text;

      // ルールベースの修正を追加
      ruleBasedResult.corrections.forEach(({ rule, original }) => {
        corrections.push({
          original,
          corrected: targetText,
          description: rule.description
        });
      });

      // 文体の一貫性の警告を追加
      if (ruleBasedResult.style.style === 'mixed') {
        corrections.unshift({
          original: targetText,
          corrected: targetText,
          description: getStyleSuggestion(ruleBasedResult.style)
        });
      }

      // AI校正
      if (!openai) {
        const apiKeySet = await promptForApiKey();
        if (!apiKeySet) {
          vscode.window.showWarningMessage('APIキーが設定されていないため、AI校正はスキップされました。');
          if (corrections.length === 0) {
            updateStatusBar(false);
            return;
          }
        }
      }

      if (openai && currentProofreadingController) {
        try {
          progress.report({ message: 'AI校正を実行中...' });

          // AI校正用のコンテキスト付きテキストを準備
          let contextualText = selection ?
            `${contextBefore}【ここから校正対象】${targetText}【ここまで校正対象】${contextAfter}` :
            targetText;

          const aiResponse = await openai.chat.completions.create({
            model: vscode.workspace.getConfiguration('otak-jproofreading').get('model', 'gpt-4-turbo-preview'),
            messages: [
              {
                role: 'system',
                content: `${selection ? '選択範囲のみを校正し、前後の文脈を考慮して、' : ''}` +
                  `日本語の文章を校正し、以下のJSON形式で返してください：\n` +
                  '{\n' +
                  '  "corrected": "修正後の文章",\n' +
                  '  "reason": "修正理由の説明"\n' +
                  '}\n\n' +
                  '以下の点を考慮してください：\n' +
                  '1. 誤字脱字の修正\n' +
                  '2. 不適切な敬語の修正\n' +
                  '3. 冗長な表現の改善\n' +
                  '4. わかりづらい表現の明確化\n' +
                  `5. 文体の統一（${ruleBasedResult.style.style === "である" ? "「である」体" : "「です・ます」体"}を維持）\n` +
                  '※ 文末表現は既存の文体を維持してください'
              },
              {
                role: 'user',
                content: contextualText
              }
            ]
          }, {
            signal: currentProofreadingController.signal
          }).catch(err => {
            if (err.name === 'AbortError') {
                throw new Error('校正がキャンセルされました');
              }
            throw err;
          });

          progress.report({ message: 'AI校正結果を解析中...' });

          try {
            let responseText = aiResponse.choices[0].message.content?.trim() || '{}';

            // レスポンステキストの解析
            
            if (selection) {
              // 選択範囲のみを抽出
              const match = responseText.match(/【ここから校正対象】([\s\S]*?)【ここまで校正対象】/);
              if (match) {
                const extractedText = match[1];
                responseText = JSON.stringify({
                  corrected: extractedText,
                  reason: responseText.includes('"reason"') ?
                    JSON.parse(responseText).reason :
                    'AI による改善提案'
                });
              }
            }

            const aiSuggestions = JSON.parse(responseText);
            if (aiSuggestions.corrected && aiSuggestions.corrected !== targetText) {
              corrections.push({
                original: targetText,
                corrected: aiSuggestions.corrected,
                description: aiSuggestions.reason || 'AI による改善提案'
              });
              targetText = aiSuggestions.corrected;
            }
          } catch (parseError) {
            console.error('AI応答解析エラー:', parseError, '\nAI応答:', aiResponse.choices[0].message.content);
            throw new Error('AI応答の解析に失敗しました。不正なJSON形式です。');
          }
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : '不明なエラー';
          if (errorMessage !== '校正がキャンセルされました') {
            console.error('AI校正エラー:', error);
            vscode.window.showErrorMessage(`AI校正中にエラーが発生しました: ${errorMessage}`);
          }
          if (corrections.length === 0) {
            updateStatusBar(false);
            return;
          }
        }
      }

      progress.report({ message: '結果を生成中...' });

      try {
        // 修正がある場合は結果を表示
        if (corrections.length > 0) {
          // オリジナルと修正後のテキストを準備
          const originalContent = text;
          const modifiedContent = targetText;

          // 修正理由のコメントを追加
          const modifiedWithComments = corrections.reduce((content, corr, index) => {
            return content + `\n\n// 修正理由 ${index + 1}: ${corr.description}`;
          }, modifiedContent);

          // 一時ファイルを作成
          const originalUri = vscode.Uri.parse('untitled:原文.md');
          const modifiedUri = vscode.Uri.parse('untitled:校正後.md');

          // ドキュメントを作成
          const originalDoc = await vscode.workspace.openTextDocument(originalUri);
          const modifiedDoc = await vscode.workspace.openTextDocument(modifiedUri);

          // コンテンツを設定
          const originalEdit = new vscode.WorkspaceEdit();
          const modifiedEdit = new vscode.WorkspaceEdit();

          originalEdit.insert(originalUri, new vscode.Position(0, 0), originalContent);
          modifiedEdit.insert(modifiedUri, new vscode.Position(0, 0), modifiedWithComments);

          await vscode.workspace.applyEdit(originalEdit);
          await vscode.workspace.applyEdit(modifiedEdit);

          // VSCodeのネイティブDiffエディタで表示
          await vscode.commands.executeCommand('vscode.diff', originalUri, modifiedUri, '校正結果');

          // 修正理由一覧を表示
          const reasonsList = corrections
            .map((corr, index) => `${index + 1}. ${corr.description}`)
            .join('\n');
          
          vscode.window.showInformationMessage('修正理由一覧を表示しますか？', '表示', 'キャンセル')
            .then(selection => {
              if (selection === '表示') {
                return vscode.window.showInformationMessage(reasonsList);
              }
            });

          // 修正を適用するボタンを表示
          const apply = await vscode.window.showInformationMessage(
            '校正結果を適用しますか？',
            '適用',
            'キャンセル'
          );

          if (apply === '適用') {
            editor.edit(editBuilder => {
              const range = selection || new vscode.Range(
                editor.document.positionAt(0),
                editor.document.positionAt(text.length)
              );
              editBuilder.replace(range, targetText);
            });
          }
        } else {
          vscode.window.showInformationMessage('修正の必要はありません。');
        }
      } finally {
        updateStatusBar(false);
      }
    });
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : '不明なエラー';
    if (errorMessage !== '校正がキャンセルされました') {
      console.error('校正エラー:', error);
      vscode.window.showErrorMessage(`校正中にエラーが発生しました: ${errorMessage}`);
    }
  } finally {
    updateStatusBar(false);
  }
}

// HTMLエスケープ
function escapeHtml(unsafe: string): string {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

export function deactivate() {
  if (statusBarItem) {
    statusBarItem.dispose();
  }
  currentCancellationTokenSource?.dispose();
}
