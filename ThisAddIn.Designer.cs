namespace PPTWebBrowserAddIn
{
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.RequiredTemplates.ThisAddIn", "17.0.0.0")]
    public partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        internal Microsoft.Office.Interop.PowerPoint.Application Application;
        internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;

        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public ThisAddIn(global::Microsoft.Office.Tools.Factory factory, global::System.IServiceProvider serviceProvider) 
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
        }

        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.RequiredTemplates.ThisAddIn", "17.0.0.0")]
        protected override void Initialize()
        {
            base.Initialize();
            
            // Only retrieve host Application explicitly. CustomTaskPanes field is populated automatically by the runtime.
            this.Application = this.GetHostItem<Microsoft.Office.Interop.PowerPoint.Application>(typeof(Microsoft.Office.Interop.PowerPoint.Application), "Application");
            
            Globals.ThisAddIn = this;
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.RequiredTemplates.ThisAddIn", "17.0.0.0")]
        protected override void FinishInitialization()
        {
            base.FinishInitialization();
        }

        // Return our custom Ribbon instance to VSTO runtime so it registers the GUI tab
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }
    }

    internal static partial class Globals
    {
        private static ThisAddIn _ThisAddIn;

        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set
            {
                if (_ThisAddIn == null)
                    _ThisAddIn = value;
                else
                    throw new System.NotSupportedException();
            }
        }
    }
}
