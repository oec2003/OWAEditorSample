using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Ican.YunPan.Services;
using Ican.YunPan.Services.GridFS;
using System.Text;
using System.Configuration;
using System.IO;

namespace OWAEditorWeb
{
    public partial class Index : System.Web.UI.Page
    {
        IDFSHandle _h = new GridFSHandle();
        string _fileId;
        string _fileSavePath = ConfigurationManager.AppSettings["FileSavePath"];
        CookieHelper _cookieHelper = new CookieHelper();

        protected void Page_Load(object sender, EventArgs e)
        {
            btnB.Attributes.Add("onclick", "this.form.target='_blank'");
            btnS.Attributes.Add("onclick", "this.form.target=''");
            btnE.Attributes.Add("onclick", "this.form.target=''");
        }

        protected void btnE_Click(object sender, EventArgs e)
        {
            _cookieHelper.GetCookieValue("FileId", out _fileId);
            if (string.IsNullOrEmpty(_fileId))
            {
                //_fileId = ConfigurationManager.AppSettings["FileId"];
                _fileId = txtId.Text.Trim();
            }

            if (string.IsNullOrEmpty(_fileId))
            {
                AlertMsg("fileidempty", "请填写文件id！");
            }

            GridFSBriefInfo info = _h.GetFileInfo(_fileId);

            StringBuilder sb = new StringBuilder();
            sb.Append("文件Id：" + info.FileId + "<br/>");
            sb.Append("文件名：" + info.FileName + "<br/>");
            sb.Append("文件大小：" + info.Length + "<br/>");
            sb.Append("文件MD5：" + info.Md5 + "<br/>");
            sb.Append("<hr/>");

            litList.Text = sb.ToString();

            var filename = info.FileName;
            Stream gridfsStream = _h.GetFile(_fileId);
            bool isExist=StreamToFile(gridfsStream, filename);
            if(!isExist)
            {
                AlertMsg("fileischeckout", "文件已被签出，签入后才能再次签出");
            }
            else
            {
                AlertMsg("checkout", "签出成功");
            }
        }

        protected void btnB_Click(object sender, EventArgs e)
        {
            _cookieHelper.GetCookieValue("FileId", out _fileId);
            if (string.IsNullOrEmpty(_fileId))
            {
                //_fileId = ConfigurationManager.AppSettings["FileId"];
                _fileId = txtId.Text.Trim();
            }

            if (string.IsNullOrEmpty(_fileId))
            {
                AlertMsg("fileidempty", "请填写文件id！");
            }

            GridFSBriefInfo info = _h.GetFileInfo(_fileId);
            var fileName = info.FileName;
            string path = Path.Combine(_fileSavePath, fileName);
            if (!File.Exists(path))
            {
                AlertMsg("edit","文件需要签出才能进行编辑！");
            }
            else
            {
                string url =GetEditorUrl(fileName);
                if(string.IsNullOrEmpty(url))
                {
                    AlertMsg("filetypeerror", "文件类型不支持编辑！");
                }
                else
                {
                    Response.Redirect(string.Format(url, fileName));
                }
            }
        }

        protected void btnS_Click(object sender, EventArgs e)
        {
            _cookieHelper.GetCookieValue("FileId", out _fileId);
            if (string.IsNullOrEmpty(_fileId))
            {
                //_fileId = ConfigurationManager.AppSettings["FileId"];
                _fileId = txtId.Text.Trim();
            }

            if (string.IsNullOrEmpty(_fileId))
            {
                AlertMsg("fileidempty", "请填写文件id！");
            }

            GridFSBriefInfo info = _h.GetFileInfo(_fileId);
            var filename = info.FileName;
            if(FileToGridfs(filename))
            {
                AlertMsg("checkin", "签入成功！");
            }
            else
            {
                AlertMsg("nofile", "没有可签入的文件！");
            }
        }

        public bool StreamToFile(Stream stream, string fileName)
        {
            string path = Path.Combine(_fileSavePath, fileName);

            if (File.Exists(path)) return false;

            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            stream.Seek(0, SeekOrigin.Begin);

            // 把 byte[] 写入文件 
            using (FileStream fs = new FileStream(path, FileMode.Create))
            {
                using (BinaryWriter bw = new BinaryWriter(fs))
                {
                    bw.Write(bytes);
                }
            }
            return true;
        }

        public bool FileToGridfs(string fileName)
        {
            string path = Path.Combine(_fileSavePath, fileName);

            if (!File.Exists(path))
            {
                return false;
            }
            // 把 byte[] 写入文件 
            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
               GridFSBriefInfo info= _h.AddFile(fs, fileName);
               _cookieHelper.CreateCookie("FileId", info.FileId);
            }

            File.Delete(path);
            return true;
        }

        private void AlertMsg(string key, string msg)
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), key,
                "<script>alert('" + msg + "');</script>");
        }

        private string GetEditorUrl(string fileName)
        {
            string wordUrl = ConfigurationManager.AppSettings["WordUrl"];
            string excelUrl = ConfigurationManager.AppSettings["ExcelUrl"];
            string pptUrl = ConfigurationManager.AppSettings["PPTUrl"];

            string result = "";
            if(fileName.ToLower().EndsWith(".doc") || fileName.ToLower().EndsWith(".docx"))
            {
                result = string.Format(wordUrl, fileName);
            }
            else if (fileName.ToLower().EndsWith(".xls") || fileName.ToLower().EndsWith(".xlsx"))
            {
                result = string.Format(excelUrl, fileName);
            }
            else if (fileName.ToLower().EndsWith(".ppt") || fileName.ToLower().EndsWith(".pptx"))
            {
                result = string.Format(pptUrl, fileName);
            }
            return result;
        }

        protected void btnC_Click(object sender, EventArgs e)
        {
            _cookieHelper.DeleteCookie("FileId");
        }
    }
}