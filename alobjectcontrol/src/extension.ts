// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import { getLastRealObjNo, clearCache } from './data/get-data';
import { clearExpiredReservedObjects, reserveId } from './data/post-data';
import { createConfigFile } from './data/helper';

var timer: NodeJS.Timeout | undefined;
// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export async function activate(context: vscode.ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "alobjectcontrol" is now active!');

	// The command has been defined in the package.json file
	// Now provide the implementation of the command with registerCommand
	// The commandId parameter must match the command field in package.json

	const completionProvider = vscode.languages.registerCompletionItemProvider({ language: 'al' }, {
		async provideCompletionItems(document, position) {
			const line = document.lineAt(position.line).text.substring(0, position.character);

			const match = line.match(/^\s*(table|tableextension|page|pageextension|report|reportextension|codeunit|xmlport|query|enum|enumextension)(?:\s+\d+)?\s*$/i);
			if (!match) {
				return undefined;
			}
			const nextId = await getLastRealObjNo(context, match[1].toString().trimEnd());

			const item = new vscode.CompletionItem(
				String(nextId),
				vscode.CompletionItemKind.Value
			);

			item.insertText = String(nextId);
			item.detail = "Next available AL object ID";

			item.command = {
				command: 'alobjectcontrol.completionSelected',
				title: 'AL Object Control: Completion selected',
				arguments: [nextId, match[1].toString().trimEnd()]
			};

			return [item];
		}
	});

	const clearTokenDisposable = vscode.commands.registerCommand('alobjectcontrol.clearCache', () => {
		// The code you place here will be executed every time your command is executed
		clearCache(context);
	});
	const createConfigFileDisposable = vscode.commands.registerCommand('alobjectcontrol.createConfigFile', () => {
		// The code you place here will be executed every time your command is executed
		createConfigFile();
	});
	const selectedCompletionDisposable = vscode.commands.registerCommand('alobjectcontrol.completionSelected', (nextId: number, objectType: string) => {
		// When my completion item is selected
		console.log('alobjectcontrol.completionSelected executed');
		Promise.all(
			[
				reserveId(context, nextId, objectType)
			]
		).then(() => {
			console.log('Posting successfull!');
		}, (reject) => {
			console.error(reject);
		});
	});

	timer = setInterval(() => {
		clearExpiredReservedObjects(context);
		console.log('Scheduled task executed!');
	}, 900000); // 15min interval

	context.subscriptions.push(clearTokenDisposable, completionProvider, selectedCompletionDisposable, createConfigFileDisposable);
}

// This method is called when your extension is deactivated
export function deactivate() {
	if (timer !== undefined) {
		clearInterval(timer);
	}
}
