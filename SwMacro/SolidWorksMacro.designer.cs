//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:2.0.50727.5485
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace smartZoom.csproj {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Contract.EntryPointAttribute(0)]
    public sealed partial class SolidWorksMacro : SolidWorks.OuterDispProxy {
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        public SolidWorksMacro(System.AddIn.Contract.Collections.IRemoteArgumentArrayContract initArgs) : 
                base(initArgs) {
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        protected override string PrimaryCookie {
            get {
                return "SolidWorksAddinApplication";
            }
        }
        
        public event System.EventHandler Startup;
        
        public event System.EventHandler Shutdown;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        protected override void OnStartup() {
            base.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            base.FinishInitialization();
            this.InternalStartup();
            if ((this.Startup != null)) {
                this.Startup(this, System.EventArgs.Empty);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        protected override void OnShutdown() {
            if ((this.Shutdown != null)) {
                this.Shutdown(this, System.EventArgs.Empty);
            }
            base.OnShutdown();
        }
    }
}
