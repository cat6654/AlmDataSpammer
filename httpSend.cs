using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace httpWebSend
{
    class httpSend
    {
        public string UserAgent = "Opera/9.80 (Windows NT 5.1; U; ru) Presto/2.2.15 Version/10.00";
        public string htmlText;

        public int pageStatus;

        public Uri RequestUri;
        public Encoding Charset;
        public Stream pageStream;
        public Stream fileStream;
        public WebResponse GetRes;
        public CookieContainer cookieJar;

        private string Url;
        private string Query;
        private bool findCaptcha;

        public httpSend(string charset)
        {
            this.cookieJar = new CookieContainer();
            switch (charset)
            {
                case "utf8":
                    this.Charset = Encoding.UTF8;
                    break;
                default:
                    this.Charset = Encoding.GetEncoding(1251);
                    break;
            }
        }

        public httpSend()
        {
            this.cookieJar = new CookieContainer();
            this.Charset = Encoding.GetEncoding(1251);
        }

        public void Post(string url, string query)
        {
            this.Url = url;
            this.Query = query;

            sendPost();
        }

        public void Post(string url, string query, bool findCaptcha)
        {
            this.Url = url;
            this.Query = query;
            this.findCaptcha = findCaptcha;

            sendPost();
        }

        public void Get(string url)
        {
            this.Url = url;
            sendGet();
            
        }

        private void createUri()
        {
            //this.uri = new Uri(this.Url);
        }

        private void sendGet()
        {
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(this.Url);
            request.ContentType = "application/xml";
            request.Accept = "application/xml";
            request.CookieContainer = cookieJar;
            this.GetRes = request.GetResponse();
            this.htmlText = new StreamReader(this.pageStream = this.GetRes.GetResponseStream()).ReadToEnd();
            this.RequestUri = request.RequestUri;
        }

        private void sendPost()
        {
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(this.Url);
            request.Method = "POST";
            request.KeepAlive = true;
            request.UserAgent = this.UserAgent;
            //request.Accept = "text/html, application/xml;q=0.9, application/xhtml+xml, image/png, image/jpeg, image/gif, image/x-xbitmap, */*;q=0.1";

            //request.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.9,en;q=0.8");
            //request.Headers.Add(HttpRequestHeader.AcceptEncoding, "deflate, gzip, x-gzip, identity, *;q=0");
            //request.Headers.Add(HttpRequestHeader.AcceptCharset, "iso-8859-1, utf-8, utf-16, *;q=0.1");
            //request.Headers.Add(HttpRequestHeader.Te, "deflate, gzip, chunked, identity, trailers");

            request.Timeout = 120000;
            request.ContentType = "application/xml";
            request.Accept = "application/xml";
            request.CookieContainer = this.cookieJar;
            
            byte[] sentData = this.Charset.GetBytes(this.Query);

            request.ContentLength = sentData.Length;
            System.IO.Stream sendStream = request.GetRequestStream();
            sendStream.Write(sentData, 0, sentData.Length);
            sendStream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Encoding _responseEncoding = Encoding.GetEncoding(response.CharacterSet);
            this.htmlText = new StreamReader(this.pageStream = response.GetResponseStream(), _responseEncoding).ReadToEnd();
            this.RequestUri = request.RequestUri;
        }

        public void getFile(string url)
        {
            //getFileStream(url);
        }

        public string fileUrl = "";
        public void getFile(string url, string filename)
        {
            this.fileUrl = url;
            //this.NewThread(getFileStream);
            Thread th = new Thread(new ThreadStart(getFileStream));
            th.Start();
            if (this.pageStatus == 200)
            {
                using (Stream file = File.OpenWrite(filename))
                {
                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    while ((len = this.fileStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, len);
                    }
                }
            }
        }

        private void getFileStream()
        {
            int iteration = 0;
            do
            {
                try
                {
                    WebRequest request = WebRequest.Create(this.fileUrl);
                    request.Credentials = CredentialCache.DefaultCredentials;
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    this.pageStatus = response.StatusCode.GetHashCode();
                    this.fileStream = response.GetResponseStream();
                }
                catch { }
                iteration++;
            } while (this.pageStatus != 200 || iteration >= 10);
        }

        private byte[] StreamToByte(Stream NonSeekableStream)
        {
            MemoryStream ms = new MemoryStream();
            byte[] buffer = new byte[1024];
            int bytes;
            while ((bytes = NonSeekableStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                ms.Write(buffer, 0, bytes);
            }

            return ms.ToArray();
        }
    }
}
