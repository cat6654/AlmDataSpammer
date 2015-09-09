using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;


namespace WindowsFormsApplication5
{
   
        public class WebGateCredential
        {
            public static void SetProxyCredential(ProxyCredential proxyCredential)
            {
                try
                {
                    SetProxyCredential(proxyCredential.serverURL, proxyCredential.port, proxyCredential.domain, proxyCredential.userName, proxyCredential.password, proxyCredential.useProxy);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            public static void SetProxyCredential(string serverURL, int port, string domain, string userName, string password, int useProxy)
            {
                try
                {
                    RegistryKey registryRootKey = Registry.CurrentUser;
                    RegistryKey registryKey = registryRootKey.OpenSubKey(@"Software\Mercury Interactive\TestDirector\Web", true);
                    switch (useProxy)
                    {
                        case 0:
                            registryKey.SetValue("Proxy", "");
                            registryKey.SetValue("ProxyPort", "");
                            registryKey.SetValue("ByPathProxy", "Y");
                            registryKey.SetValue("UseIEProxy", "N");
                            registryKey.SetValue("UseQTProxy", "N");
                            break;
                        case 1:
                            registryKey.SetValue("Proxy", "");
                            registryKey.SetValue("ProxyPort", "");
                            registryKey.SetValue("ByPathProxy", "N");
                            registryKey.SetValue("UseIEProxy", "Y");
                            registryKey.SetValue("UseQTProxy", "N");
                            break;
                        case 2:
                            registryKey.SetValue("Proxy", serverURL);
                            registryKey.SetValue("ProxyPort", port.ToString());
                            registryKey.SetValue("ByPathProxy", "N");
                            registryKey.SetValue("UseIEProxy", "N");
                            registryKey.SetValue("UseQTProxy", "Y");
                            break;
                        default:
                            throw new Exception("Unknow UseProxy value.0 for no proxy in use. 1 for use IE proxy. 2 for use specified proxy. Any other value is not acceptable.");
                    }
                    if (domain != "") userName = domain + "\\" + userName;
                    registryKey.SetValue("ProxyAuthUser", userName);
                    registryKey.SetValue("ProxyAuthPswrd", Encrypt(password));
                    registryKey.SetValue("ProxyAuthDomain", "");
                    registryKey.SetValue("IISAuthUser", "");
                    registryKey.SetValue("IISAuthPswrd", "");
                    registryKey.SetValue("IISAuthDomain", "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            public static void SetWebServerCredential(WebServerCredential webServerCredential)
            {
                try
                {
                    SetWebServerCredential(webServerCredential.domain, webServerCredential.userName, webServerCredential.password);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            public static void SetWebServerCredential(string domain, string userName, string password)
            {
                try
                {
                    RegistryKey registryRootKey = Registry.CurrentUser;
                    RegistryKey registryKey = registryRootKey.OpenSubKey(@"Software\Mercury Interactive\TestDirector\Web", true);
                    if (domain != "") userName = domain + "\\" + userName;
                    registryKey.SetValue("Proxy", "");
                    registryKey.SetValue("ProxyPort", "");
                    registryKey.SetValue("ByPathProxy", "N");
                    registryKey.SetValue("UseIEProxy", "Y");
                    registryKey.SetValue("UseQTProxy", "N");
                    registryKey.SetValue("IISAuthUser", userName);
                    registryKey.SetValue("IISAuthPswrd", Encrypt(password));
                    registryKey.SetValue("IISAuthDomain", "");
                    registryKey.SetValue("ProxyAuthUser", "");
                    registryKey.SetValue("ProxyAuthPswrd", "");
                    registryKey.SetValue("ProxyAuthDomain", "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            public static void ClearWebServerCredential()
            {
                try
                {
                    RegistryKey registryRootKey = Registry.CurrentUser;
                    RegistryKey registryKey = registryRootKey.OpenSubKey(@"Software\Mercury Interactive\TestDirector\Web", true);
                    registryKey.SetValue("Proxy", "");
                    registryKey.SetValue("ProxyPort", "");
                    registryKey.SetValue("ByPathProxy", "N");
                    registryKey.SetValue("UseIEProxy", "N");
                    registryKey.SetValue("UseQTProxy", "N");
                    registryKey.SetValue("IISAuthUser", "");
                    registryKey.SetValue("IISAuthPswrd", "");
                    registryKey.SetValue("IISAuthDomain", "");
                    registryKey.SetValue("ProxyAuthUser", "");
                    registryKey.SetValue("ProxyAuthPswrd", "");
                    registryKey.SetValue("ProxyAuthDomain", "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            public static string Encrypt(string source)
            {
                try
                {
                    if (source == "") return "";
                    string pwd = "SmolkaWasHereMonSher";

                    int sourceSize = source.Length + 1;
                    int pwdSize = pwd.Length + 1;
                    int x = 0;
                    char[] temp = new char[source.Length];
                    for (int i = 0; i <= sourceSize - 3; i++)
                    {
                        temp[i] = Convert.ToChar(Convert.ToInt16(source[i]) + Convert.ToInt16(pwd[x]));
                        x++;
                        if (x == (pwdSize - 1)) x = 0;
                    }
                    temp[source.Length - 1] = source[source.Length - 1];

                    byte[] bytes = new byte[temp.Length];
                    for (int index = 0; index < temp.Length; index++)
                    {
                        bytes[index] = (byte)temp[index];
                    }
                    string result = Encoding.Default.GetString(bytes);
                    return result;
                }
                catch (Exception)
                {
                    throw;
                }
            }

        }

    }

