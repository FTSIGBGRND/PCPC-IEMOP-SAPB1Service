using System;
using System.ComponentModel;

namespace SAPB1iService
{
    partial class SAPB1Service
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.runInitialization = new System.ComponentModel.BackgroundWorker();
            this.runProcService = new System.ComponentModel.BackgroundWorker();
            // 
            // runInitialization
            // 
            this.runInitialization.DoWork += new System.ComponentModel.DoWorkEventHandler(this.runInitialization_DoWork);
            this.runInitialization.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.runInitialization_RunWorkerCompleted);
            // 
            // runProcService
            // 
            this.runProcService.DoWork += new System.ComponentModel.DoWorkEventHandler(this.runProcService_DoWork);
            this.runProcService.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.RunProcService_RunWorkerCompleted);
            // 
            // SAPB1Service
            // 
            this.ServiceName = "SAPB1Service";

        }


        #endregion

        private System.ComponentModel.BackgroundWorker runInitialization;
        private BackgroundWorker runProcService;
    }
}
