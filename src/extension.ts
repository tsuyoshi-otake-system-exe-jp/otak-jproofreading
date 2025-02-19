// 先ほどの内容と同じですが、以下の修正を加えます：

// 1. asyncでwithProgressブロックを正しく構造化
// 2. エラーハンドリングの改善
// 3. 非同期処理の最適化

// 既存のインポートと型定義はそのまま維持

import * as vscode from 'vscode';
import OpenAI from 'openai';
import { HttpsProxyAgent } from 'https-proxy-agent';
import * as https from 'https';

// インターフェース定義（既存のものを維持）
interface ProofreadingRule {
  pattern: RegExp;
  replacement: string | ((match: string, ...args: string[]) => string);
  description: string;
}

// 定数定義（既存のものを維持）
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

// 型定義（既存のものを維持）
interface CorrectionItem {
  original: string;
  corrected: string;
  description: string;
}

interface RuleCorrection {
  rule: ProofreadingRule;
  original: string;
}

interface ProofreadingState {
  text: string;
  targetText: string;
  corrections: CorrectionItem[];
}

interface DisplayResultsParams {
  originalText: string;
  modifiedText: string;
  corrections: CorrectionItem[];
  editor: vscode.TextEditor;
  selection?: vscode.Selection;
}

interface AIResult {
  text: string;
  response: string | null;
  correction: { corrected: string; reason: string; } | null;
}

interface ProofreadingContext {
  editor: vscode.TextEditor;
  fullText: string;
  targetText: string;
  selection?: vscode.Selection;
  contextBefore: string;
  contextAfter: string;
  openai: OpenAI | undefined;
  progress: vscode.Progress<{ message?: string }>;
  controller: AbortController;
}

// エラーハンドリング用の型
class ProofreadingError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ProofreadingError';
  }
}

// ユーティリティ関数
function assertProofreadingContext(context: Partial<ProofreadingContext>): asserts context is ProofreadingContext {
  if (!context.editor || !context.fullText || !context.targetText || !context.progress || !context.controller) {
    throw new ProofreadingError('必要なコンテキストが不足しています');
  }
}

// 既存の関数（変更なし）
function detectWritingStyle(text: string): {
  style: 'ですます' | 'である' | 'mixed' | 'unknown';
  details: { dearu: number; desuMasu: number; };
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

// 既存の関数（変更なし）
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

// プロキシ設定を取得する関数（既存のまま）
function getProxySettings(): { httpsProxy: string | undefined } {
  const config = vscode.workspace.getConfiguration('otak-jproofreading');
  const configuredProxy = config.get<string>('proxy');

  if (configuredProxy) {
    return { httpsProxy: configuredProxy };
  }

  const httpSettings = vscode.workspace.getConfiguration('http');
  const proxy = httpSettings.get<string>('proxy');

  if (proxy) {
    return { httpsProxy: proxy };
  }

  return {
    httpsProxy: process.env.HTTPS_PROXY || process.env.HTTP_PROXY
  };
}

// OpenAIクライアントを初期化する関数（既存のまま）
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

  if (httpsProxy) {
    const agent = new HttpsProxyAgent(httpsProxy);
    configuration.httpAgent = agent;
    configuration.httpsAgent = agent;
  }

  return new OpenAI(configuration);
}

// APIキーの設定を促す関数（既存のまま）
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

// ステータスバーの表示を更新（既存のまま）
function updateStatusBar(isChecking: boolean = false): void {
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

// 校正をキャンセル（既存のまま）
function cancelProofreading(): void {
  currentProofreadingController?.abort();
  currentCancellationTokenSource?.cancel();
  updateStatusBar(false);
}

// activate関数（既存のまま）
export function activate(context: vscode.ExtensionContext) {
  statusBarItem = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Right,
    100
  );
  updateStatusBar();
  statusBarItem.show();

  openai = initializeOpenAIClient();

  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration(e => {
      if (e.affectsConfiguration('otak-jproofreading') || e.affectsConfiguration('http')) {
        openai = initializeOpenAIClient();
      }
    })
  );

  const cancelProofreadingDisposable = vscode.commands.registerCommand(
    'otak-jproofreading.cancelProofreading',
    cancelProofreading
  );

  const checkDocumentDisposable = vscode.commands.registerCommand(
    'otak-jproofreading.checkDocument',
    async () => {
      if (isProofreading) {
        return cancelProofreading();
      }
      const editor = vscode.window.activeTextEditor;
      if (!editor) {
        vscode.window.showWarningMessage('テキストエディタがアクティブではありません。');
        return;
      }

      const document = editor.document;
      const text = document.getText();
      await proofreadText(text, editor);
    }
  );

  const checkSelectionDisposable = vscode.commands.registerCommand(
    'otak-jproofreading.checkSelection',
    async () => {
      if (isProofreading) {
        return cancelProofreading();
      }
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
    }
  );

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

    await vscode.window.withProgress({
      location: vscode.ProgressLocation.Notification,
      title: "校正を開始します",
      cancellable: true
    }, async (progress, token) => {
      token.onCancellationRequested(() => {
        cancelProofreading();
      });

      const state: ProofreadingState = {
        text,
        targetText: text,
        corrections: []
      };

      const fullText = editor.document.getText();
      let contextBefore = '';
      let contextAfter = '';
      
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

      const styleAnalysis = await detectWritingStyle(fullText);

      progress.report({ message: 'ルールベース校正を実行中...' });

      let currentText = state.targetText;
      const ruleCorrections: RuleCorrection[] = [];

      for (const rule of PROOFREADING_RULES) {
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
      }

      state.targetText = currentText;

      ruleCorrections.forEach(({ rule, original }) => {
        state.corrections.push({
          original,
          corrected: state.targetText,
          description: rule.description
        });
      });

      if (styleAnalysis.style === 'mixed') {
        state.corrections.unshift({
          original: state.targetText,
          corrected: state.targetText,
          description: getStyleSuggestion(styleAnalysis)
        });
      }

      if (!openai && currentProofreadingController) {
        const apiKeySet = await promptForApiKey();
        if (!apiKeySet) {
          if (state.corrections.length === 0) {
            vscode.window.showWarningMessage('APIキーが設定されていないため、AI校正はスキップされました。');
            return;
          }
        }
      }

      if (openai && currentProofreadingController) {
        progress.report({ message: 'AI校正を実行中...' });

        const context: ProofreadingContext = {
          editor,
          fullText,
          targetText: state.targetText,
          selection,
          contextBefore,
          contextAfter,
          openai,
          progress,
          controller: currentProofreadingController
        };

        assertProofreadingContext(context);

        const aiResult = await processAIResponse(context);

        if (aiResult.correction) {
          state.corrections.push({
            original: state.targetText,
            corrected: aiResult.correction.corrected,
            description: aiResult.correction.reason
          });
          state.targetText = aiResult.correction.corrected;
        }
      }

      if (state.corrections.length > 0) {
        await displayResults({
          originalText: state.text,
          modifiedText: state.targetText,
          corrections: state.corrections,
          editor,
          selection
        });
      } else {
        vscode.window.showInformationMessage('修正の必要はありません。');
      }
    });
  } catch (error) {
    handleError(error);
  } finally {
    updateStatusBar(false);
  }
}

// AI応答の処理
async function processAIResponse(context: ProofreadingContext): Promise<AIResult> {
  const { openai, targetText, selection, contextBefore, contextAfter, controller } = context;

  if (!openai) {
    return { text: targetText, response: null, correction: null };
  }

  try {
    const aiResponse = await openai.chat.completions.create({
      model: vscode.workspace.getConfiguration('otak-jproofreading').get('model', 'gpt-4-turbo-preview'),
      messages: [
        {
          role: 'system',
          content: createSystemPrompt(context)
        },
        {
          role: 'user',
          content: selection ?
            `${contextBefore}【ここから校正対象】${targetText}【ここまで校正対象】${contextAfter}` :
            targetText
        }
      ]
    }, {
      signal: controller.signal
    });

    const responseText = aiResponse.choices[0].message.content?.trim() || '{}';
    return {
      text: targetText,
      response: responseText,
      correction: parseAIResponse(responseText, selection)
    };
  } catch (error) {
    console.error('AI校正エラー:', error);
    return { text: targetText, response: null, correction: null };
  }
}

// システムプロンプトの生成
function createSystemPrompt(context: ProofreadingContext): string {
  return `${context.selection ? '選択範囲のみを校正し、前後の文脈を考慮して、' : ''}` +
    '日本語の文章を校正し、以下のJSON形式で返してください：\n' +
    '{\n' +
    '  "corrected": "修正後の文章",\n' +
    '  "reason": "修正理由の説明"\n' +
    '}\n\n' +
    '以下の点を考慮してください：\n' +
    '1. 誤字脱字の修正\n' +
    '2. 不適切な敬語の修正\n' +
    '3. 冗長な表現の改善\n' +
    '4. わかりやすい表現への変換\n' +
    '5. 文体の統一\n' +
    '※ 文末表現は既存の文体を維持してください';
}

// AI応答の解析
function parseAIResponse(
  responseText: string,
  selection?: vscode.Selection
): { corrected: string; reason: string } | null {
  try {
    if (selection) {
      const match = responseText.match(/【ここから校正対象】([\s\S]*?)【ここまで校正対象】/);
      if (match?.[1]) {
        return {
          corrected: match[1],
          reason: responseText.includes('"reason"') ?
            JSON.parse(responseText).reason :
            'AI による改善提案'
        };
      }
    }

    try {
      const parsed = JSON.parse(responseText);
      if (parsed.corrected && parsed.reason) {
        return parsed;
      }
    } catch (parseError) {
      console.error('JSON解析エラー:', parseError);
    }

    return null;
  } catch (error) {
    console.error('AI応答解析エラー:', error);
    return null;
  }
}

// エラー処理
function handleError(error: unknown): void {
  if (error instanceof ProofreadingError) {
    void vscode.window.showErrorMessage(`校正エラー: ${error.message}`);
    return;
  }

  const errorMessage = error instanceof Error ? error.message : '不明なエラー';
  if (errorMessage !== '校正がキャンセルされました') {
    console.error('校正エラー:', error);
    void vscode.window.showErrorMessage(`校正中にエラーが発生しました: ${errorMessage}`);
  }
}

// 結果表示処理
async function displayResults({
  originalText,
  modifiedText,
  corrections,
  editor,
  selection
}: DisplayResultsParams): Promise<void> {
  const modifiedWithComments = corrections.reduce(
    (content: string, corr: CorrectionItem, index: number) =>
      content + `\n\n// 修正理由 ${index + 1}: ${corr.description}`,
    modifiedText
  );

  const originalUri = vscode.Uri.parse('untitled:原文.md');
  const modifiedUri = vscode.Uri.parse('untitled:校正後.md');

  const originalEdit = new vscode.WorkspaceEdit();
  const modifiedEdit = new vscode.WorkspaceEdit();

  originalEdit.insert(originalUri, new vscode.Position(0, 0), originalText);
  modifiedEdit.insert(modifiedUri, new vscode.Position(0, 0), modifiedWithComments);

  await vscode.workspace.applyEdit(originalEdit);
  await vscode.workspace.applyEdit(modifiedEdit);

  await vscode.commands.executeCommand('vscode.diff', originalUri, modifiedUri, '校正結果');
  await showCorrectionReasons(corrections);
  await promptToApplyChanges(editor, modifiedText, selection);
}

// 修正理由の表示
async function showCorrectionReasons(corrections: CorrectionItem[]): Promise<void> {
  const reasonsList = corrections
    .map((corr, index) => `${index + 1}. ${corr.description}`)
    .join('\n');

  void vscode.window.showInformationMessage(reasonsList, { modal: true });
}

// 修正の適用
async function promptToApplyChanges(
  editor: vscode.TextEditor,
  modifiedText: string,
  selection?: vscode.Selection
): Promise<void> {
  const apply = await vscode.window.showInformationMessage(
    '校正結果を適用しますか？',
    '適用',
    'キャンセル'
  );

  if (apply === '適用') {
    await editor.edit(editBuilder => {
      const range = selection || new vscode.Range(
        editor.document.positionAt(0),
        editor.document.positionAt(modifiedText.length)
      );
      editBuilder.replace(range, modifiedText);
    });
  }
}

// deactivate関数（既存のまま）
export function deactivate() {
  if (statusBarItem) {
    statusBarItem.dispose();
  }
  currentCancellationTokenSource?.dispose();
}
