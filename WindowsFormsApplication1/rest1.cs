using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Threading;
using System.Web;
using System.Text.RegularExpressions;


namespace WindowsFormsApplication5
{
    public class rest1
    {
        public string myheader;
        public string url;
        public string xml;
        public string backstr;
        public string myheaders;

        public void auth(string url, string xml)
        {
              //  xml = "<?xml version='1.0' encoding='UTF-8'?><alm-authentication><user>alex_k</user><password></password></alm-authentication>";
              //  url = "http://qcc-jabba.vlab.lohika.com:8080/qcbin/authentication-point/alm-authenticate";
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);

                byte[] requestBytes = System.Text.Encoding.UTF8.GetBytes(xml.ToString());
                req.Method = "POST";
                req.ContentType = "application/xml";
                req.Accept = "application/xml";
                req.KeepAlive = true;
                req.AllowAutoRedirect = true;

                req.ContentLength = requestBytes.Length;
                Stream requestStream = req.GetRequestStream();
                requestStream.Write(requestBytes, 0, requestBytes.Length);
                requestStream.Close();

                HttpWebResponse res = (HttpWebResponse)req.GetResponse();
                StreamReader sr = new StreamReader(res.GetResponseStream(), System.Text.Encoding.Default);
                backstr = sr.ReadToEnd();

               // backstr = res.Headers.Get("Set-Cookie");
                myheader = res.Headers.Get("Set-Cookie");

                //textBox1.Text = backstr;

                sr.Close();
                res.Close();

           
            
        }

        public void get(string url)
        {
            

                
               // url = "http://qcc-jabba.vlab.lohika.com:8080/qcbin/rest/domains/ALEX_K/projects/ora/defects/1";
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Accept = "application/xml";
                req.KeepAlive = true;
                req.AllowAutoRedirect = true;
                req.Headers.Add("Cookie", myheaders);
                //req.Headers.Add("Cookie", myheader);
                

                HttpWebResponse res = (HttpWebResponse)req.GetResponse();
                StreamReader sr = new StreamReader(res.GetResponseStream(), System.Text.Encoding.Default);
                backstr = sr.ReadToEnd();
               
                
                sr.Close();
                res.Close();

                        
        }
        
        public void post(string url, string xml)
        {
            //  xml = "<?xml version='1.0' encoding='UTF-8'?><alm-authentication><user>alex_k</user><password></password></alm-authentication>";
            //  url = "http://qcc-jabba.vlab.lohika.com:8080/qcbin/authentication-point/alm-authenticate";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);

            if (xml != null)
            {
                byte[] requestBytes = System.Text.Encoding.UTF8.GetBytes(xml.ToString());

                req.Method = "POST";
                req.ContentType = "application/xml";
                req.Accept = "application/xml";
                req.KeepAlive = true;
                req.AllowAutoRedirect = true;
                req.Headers.Add("Cookie", myheaders);
                //req.Headers.Add("Cookie", myheader);

                req.ContentLength = requestBytes.Length;
                Stream requestStream = req.GetRequestStream();
                requestStream.Write(requestBytes, 0, requestBytes.Length);
                requestStream.Close();

                HttpWebResponse res = (HttpWebResponse)req.GetResponse();
                StreamReader sr = new StreamReader(res.GetResponseStream(), System.Text.Encoding.Default);
                backstr = sr.ReadToEnd();

                sr.Close();
                res.Close();
            }

            req.Method = "POST";
            req.ContentType = "text/plain";
            req.KeepAlive = true;
            req.AllowAutoRedirect = true;
            req.Headers.Add("Cookie", myheader);

            HttpWebResponse res2 = (HttpWebResponse)req.GetResponse();
            StreamReader sr2 = new StreamReader(res2.GetResponseStream(), System.Text.Encoding.Default);
            backstr = sr2.ReadToEnd();
            myheaders = res2.Headers.Get("Set-Cookie");

            sr2.Close();
            res2.Close();


        }


    }
}