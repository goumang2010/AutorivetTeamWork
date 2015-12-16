using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace AutorivetService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
            System.ServiceProcess.ServiceController s = new System.ServiceProcess.ServiceController("AutorivetService");
            SetUIEnable("AutorivetService");
            s.Start();
            
        }

        private void SetUIEnable(string serviceName)
        {
            System.Management.ManagementObject wmi = new System.Management.ManagementObject(string.Format("Win32_Service.Name='{0}'", serviceName));
            System.Management.ManagementBaseObject cm = wmi.GetMethodParameters("Change");
            cm["DesktopInteract"] = true;
            System.Management.ManagementBaseObject ut = wmi.InvokeMethod("Change", cm, null);
        }
    }
}
