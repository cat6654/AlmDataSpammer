using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using HP.QC.InternalIntegration;
using System.Net;
using System.Windows.Forms;
using Mercury.TD.Client.UI.Api;
using Mercury.TD.Client.Ota.Api;
using TDAPIOLELib;
using EntitiesConstants = Mercury.TD.Client.Ota.Entities.Core.Constants;


namespace WindowsFormsApplication5
{
  public class QCCH_API
    {

        public static QCConnectivityHelper m_connectivityHelper;
        public static IConnectionManagementService connectionManagementService;
        public static INewEntityDialogService entityManagementService;
        public static IFactory myFactory;
        public static IServiceManager m_ServiceManager = (IServiceManager)m_connectivityHelper.GetServiceManager();
        public static IServiceManager ServiceManager { get; set; }

        public static void download(string serverUrl)
        {
            try
            {
                
                //m_connectivityHelper.SetQCVersionLimits(new Version(10, 0), null);
                m_connectivityHelper.AsyncOperation = false;

                //IsolatedDeployment = false -> Common
                //IsolatedDeployment = true; SharedDeployment = true -> Shared

                m_connectivityHelper.IsolatedDeployment = true;
                m_connectivityHelper.AllowLoaderDownload = true;
                m_connectivityHelper.SkipOsCheck = false;
                m_connectivityHelper.LocalNoServer = false;

               // m_connectivityHelper.CreateActivationContext();


                //m_connectivityHelper.OnInvalidSSLServerCertificate += CertificateValidationCallback;

                //m_connectivityHelper.ClientCertificates.Add(new X509Certificate(@"C:\Users\wangzhe\Desktop\Temp\To be deleted\Cert\client1.p12","W8Fwg3jT"));
                // m_connectivityHelper.ClientCertificates.Add(new X509Certificate(@"C:\Users\wangzhe\Desktop\Temp\To be deleted\Mike.cer"));

               // string flavor = "OTAAndUI";
               string flavor = "OTAOnly";

                

                ThreadStart starter = delegate { download(serverUrl, flavor); };
                if (m_connectivityHelper.AsyncOperation == false)
                {
                    starter.Invoke();
                }
                else
                {
                    Thread newThread = new Thread(starter);

                    // STA Required since QCCH does ActiveX hosting:
                    newThread.SetApartmentState(ApartmentState.STA);
                    newThread.Name = "QCCH";
                    newThread.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void initial()
        {
            m_connectivityHelper = new QCConnectivityHelper();
        }

        public static void getServiceManager()
        {
            try
            {
                
               // IServiceManager m_ServiceManager = (IServiceManager)m_connectivityHelper.GetServiceManager();
                m_connectivityHelper.CreateActivationContext();
                connectionManagementService = m_ServiceManager.GetService<IConnectionManagementService>();
                connectionManagementService.Initialize(m_ServiceManager);
                ((ISupportClientType)connectionManagementService.Site).SetClientType("Sprinter", "a44dda137c922acfd3cb511994cf01bb");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void getFactory()
        {
            try
            {
                IDictionary<string, object> fieldValues = new Dictionary<string, object>();
                fieldValues.Add(EntitiesConstants.Bug.SUMMARY, "DEFECT FROM QCCH Testign Tool - Rock&Roll");

                INewEntityDialogService newEntityDialogService = ServiceManager.GetService<INewEntityDialogService>();

                //INewEntityDialogContext context = new NewEntityDialogContext(null, true, false, fieldValues, null, null);
                //newEntityDialogService.ShowDialog<IBug>(context);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        public static void login(string userName, string password, string domain, string project)
        {
            try
            {
                connectionManagementService.Authenticate(userName, password, false);
                connectionManagementService.Login(domain, project);
                connectionManagementService.Connection.KeepConnection = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void logout()
        {
            try
            {
                connectionManagementService.Logout(true);
                connectionManagementService.DisconnectFromSite();
                //m_connectivityHelper.DestroyActivationContext();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public static void end()
        {
            try
            {
                m_connectivityHelper.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static void download(string url, string flavor, bool isSA = false)
        {
            try
            {
                if (isSA == false)
                {
                    m_connectivityHelper.DownloadComponents(url, flavor);
                }
                else
                {
                    m_connectivityHelper.DownloadSiteAdminComponents(url, flavor);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void setProxyCredential(ProxyCredential proxyCredential)
        {
            try
            {
                setProxyCredential(proxyCredential.serverURL, proxyCredential.port, proxyCredential.domain, proxyCredential.userName, proxyCredential.password);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void setProxyCredential(string serverURL, int port, string domain, string userName, string password)
        {
            try
            {
                m_connectivityHelper.ProxyServerURL = serverURL;
                m_connectivityHelper.ProxyServerPort = port;
                m_connectivityHelper.ProxyServerCredentials = new NetworkCredential(userName, password, domain);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void setWebServerCredential(WebServerCredential webServerCredential)
        {
            try
            {
                setWebServerCredential(webServerCredential.domain, webServerCredential.userName, webServerCredential.password);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void setWebServerCredential(string domain, string userName, string password)
        {
            try
            {
                m_connectivityHelper.WebServerCredentials = new NetworkCredential(userName, password, domain);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void clearAllCredential()
        {
            try
            {
                m_connectivityHelper.WebServerCredentials = null;
                m_connectivityHelper.ProxyServerCredentials = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static bool CertificateValidationCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslpolicyerrors)
        {
            // Validation logic
            return true;
        }
    }

    public class ProxyCredential
    {
        public string serverURL;
        public int port;
        public string domain;
        public string userName;
        public string password;
        public int useProxy; //0 for no proxy in use. 1 for use IE proxy. 2 for use specified proxy.

        public ProxyCredential(string serverURL, int port, string domain, string userName, string password, int useProxy)
        {
            this.serverURL = serverURL;
            this.port = port;
            this.domain = domain;
            this.userName = userName;
            this.password = password;
            this.useProxy = useProxy;
        }
    }

    public class WebServerCredential
    {
        public string domain;
        public string userName;
        public string password;

        public WebServerCredential(string domain, string userName, string password)
        {
            this.domain = domain;
            this.userName = userName;
            this.password = password;
        }


    }
}
