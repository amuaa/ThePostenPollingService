namespace ThePostenPollingService
{
    partial class ProjectInstaller
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
            this.ThePostenPollingServiceInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.ThePostenPollingService = new System.ServiceProcess.ServiceInstaller();
            // 
            // ThePostenPollingServiceInstaller
            // 
            this.ThePostenPollingServiceInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.ThePostenPollingServiceInstaller.Password = null;
            this.ThePostenPollingServiceInstaller.Username = null;
           
            // 
            // ThePostenPollingService
            // 
            this.ThePostenPollingService.ServiceName = "ThePostenPollingService";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.ThePostenPollingServiceInstaller,
            this.ThePostenPollingService});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller ThePostenPollingServiceInstaller;
        private System.ServiceProcess.ServiceInstaller ThePostenPollingService;
    }
}