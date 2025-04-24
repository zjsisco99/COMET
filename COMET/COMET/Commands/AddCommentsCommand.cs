using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Linq;
using Task = System.Threading.Tasks.Task;
using System.Threading.Tasks;

namespace COMET
{
    internal sealed class CommentMethodCommand
    {
        public const int CommandId = 0x0200;
        public static readonly Guid CommandSet = new Guid("95db4ddb-cbc1-4e37-85fe-538c13eb0bda");
        private readonly AsyncPackage _package;

        private CommentMethodCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (!(await package.GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService commandService))
                return;
            new CommentMethodCommand(package, commandService);
        }
        private async Task<ToolWindow1Control> EnsureToolWindowControlAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (COMETPackage.ToolWindowControlInstance != null)
                return COMETPackage.ToolWindowControlInstance;

            // Use the correct reference to the package
            ToolWindowPane window = await _package.ShowToolWindowAsync(typeof(ToolWindow1), 0, true, _package.DisposalToken);
            if (window == null || window.Frame == null)
                throw new NotSupportedException("Cannot create tool window");

            var control = (ToolWindow1Control)window.Content;
            COMETPackage.ToolWindowControlInstance = control;
            return control;
        }

        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            DTE2 dte = await _package.GetServiceAsync(typeof(DTE)) as DTE2;
            if (dte?.ActiveDocument == null)
                return;

            Document activeDoc = dte.ActiveDocument;
            var textDoc = (TextDocument)activeDoc.Object("TextDocument");
            EditPoint start = textDoc.StartPoint.CreateEditPoint();
            string fullText = start.GetText(textDoc.EndPoint);

            SyntaxTree tree = CSharpSyntaxTree.ParseText(fullText);
            var root = tree.GetRoot();

            TextSelection selection = (TextSelection)dte.ActiveDocument.Selection;
            int cursorLine = selection.ActivePoint.Line;
            var targetNode = root.FindToken(selection.ActivePoint.AbsoluteCharOffset).Parent;

            var node = root.DescendantNodes()
    .Where(n => n.GetLocation().GetMappedLineSpan().StartLinePosition.Line + 1 <= cursorLine &&
                n.GetLocation().GetMappedLineSpan().EndLinePosition.Line + 1 >= cursorLine)
    .OrderBy(n => n.Span.Length)
    .FirstOrDefault(n => n is MethodDeclarationSyntax ||
                         n is ConstructorDeclarationSyntax ||
                         n is PropertyDeclarationSyntax ||
                         n is ClassDeclarationSyntax ||
                         n is EnumDeclarationSyntax ||
                         n is StructDeclarationSyntax ||
                         n is NamespaceDeclarationSyntax);


            if (node == null)
                return;

            ToolWindow1Control control = await EnsureToolWindowControlAsync();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (node is MethodDeclarationSyntax method)
                control.InsertMethodComment(method);
            else if (node is ConstructorDeclarationSyntax ctor) 
                control.InsertConstructorComment(ctor);
            else if (node is PropertyDeclarationSyntax prop)
                control.InsertPropertyComment(prop);
            else if (node is ClassDeclarationSyntax cls)
                control.InsertClassComment(cls);
            else if (node is EnumDeclarationSyntax en)
                control.InsertEnumComment(en);
            else if (node is StructDeclarationSyntax st)
                control.InsertStructComment(st);
            else if (node is NamespaceDeclarationSyntax ns)
                control.InsertNamespaceComment(ns);

        }
    }
}