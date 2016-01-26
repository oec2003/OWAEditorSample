// Copyright 2014 The Authors Marx-Yu. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

using Cobalt;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using Ican.YunPan.Services;
using Ican.YunPan.Services.GridFS;

namespace WopiCobaltHost
{
    public class CobaltServer
    {
        private HttpListener m_listener;
        private string m_docsPath;
        private int m_port;

        public CobaltServer(string docsPath, int port = 8080)
        {
            m_docsPath = docsPath;
            m_port = port;
        }

        public void Start()
        {
            m_listener = new HttpListener();
            m_listener.Prefixes.Add(String.Format("http://localhost:{0}/wopi/", m_port));
            m_listener.Start();
            m_listener.BeginGetContext(ProcessRequest, m_listener);

            Console.WriteLine(@"WopiServer Started");
        }

        public void Stop()
        {
            m_listener.Stop();
        }

        private void ErrorResponse(HttpListenerContext context, string errmsg)
        {
            byte[] buffer = Encoding.UTF8.GetBytes(errmsg);
            context.Response.ContentLength64 = buffer.Length;
            context.Response.ContentType = @"application/json";
            context.Response.OutputStream.Write(buffer, 0, buffer.Length);
            context.Response.OutputStream.Close();
            context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
            context.Response.Close();
        }

        private void ProcessRequest(IAsyncResult result)
        {
            try
            {
                Console.WriteLine("start...");
                HttpListener listener = (HttpListener)result.AsyncState;
                HttpListenerContext context = listener.EndGetContext(result);
                try
                {
                    Console.WriteLine("1111...");
                    Console.WriteLine(context.Request.HttpMethod + @" " + context.Request.Url.AbsolutePath);
                    var stringarr = context.Request.Url.AbsolutePath.Split('/');
                    var access_token = context.Request.QueryString["access_token"];

                    if (stringarr.Length < 3 || access_token == null)
                    {
                        Console.WriteLine(@"Invalid request");
                        ErrorResponse(context, @"Invalid request parameter");
                        m_listener.BeginGetContext(ProcessRequest, m_listener);
                        return;
                    }

                    //todo:
                    
                    
                    //string fileId = stringarr[3];
                    var filename = stringarr[3];
                    //Stream gridfsStream = GetFileById(fileId);
                    //StreamToFile(gridfsStream, filename);


                    //use filename as session id just test, recommend use file id and lock id as session id
                    EditSession editSession = CobaltSessionManager.Instance.GetSession(filename);
                    if (editSession == null)
                    {
                        Console.WriteLine("2222...");
                        var fileExt = filename.Substring(filename.LastIndexOf('.') + 1);
                        if (fileExt.ToLower().Equals(@"xlsx"))
                            editSession = new FileSession(filename, m_docsPath + "/" + filename, @"yonggui.yu", @"yuyg", @"yonggui.yu@emacle.com", false);
                        else
                            editSession = new CobaltSession(filename, m_docsPath + "/" + filename, @"yonggui.yu", @"yuyg", @"yonggui.yu@emacle.com", false);
                        CobaltSessionManager.Instance.AddSession(editSession);
                    }
            
                    if (stringarr.Length == 4 && context.Request.HttpMethod.Equals(@"GET"))
                    {
                        Console.WriteLine("4444...");
                        //request of checkfileinfo, will be called first
                        var memoryStream = new MemoryStream();
                        var json = new DataContractJsonSerializer(typeof(WopiCheckFileInfo));
                        json.WriteObject(memoryStream, editSession.GetCheckFileInfo());
                        memoryStream.Flush();
                        memoryStream.Position = 0;
                        StreamReader streamReader = new StreamReader(memoryStream);
                        var jsonResponse = Encoding.UTF8.GetBytes(streamReader.ReadToEnd());

                        context.Response.ContentType = @"application/json";
                        context.Response.ContentLength64 = jsonResponse.Length;
                        context.Response.OutputStream.Write(jsonResponse, 0, jsonResponse.Length);
                        context.Response.Close();
                    }
                    else if (stringarr.Length == 5 && stringarr[4].Equals(@"contents"))
                    {
                        Console.WriteLine("5555...");
                        // get and put file's content, only for xlsx and pptx
                        if (context.Request.HttpMethod.Equals(@"POST"))
                        {
                            var ms = new MemoryStream();
                            context.Request.InputStream.CopyTo(ms);
                            editSession.Save(ms.ToArray());
                            context.Response.ContentLength64 = 0;
                            context.Response.ContentType = @"text/html";
                            context.Response.StatusCode = (int)HttpStatusCode.OK;
                        }
                        else
                        {
                            var content = editSession.GetFileContent();
                            context.Response.ContentType = @"application/octet-stream";
                            context.Response.ContentLength64 = content.Length;
                            context.Response.OutputStream.Write(content, 0, content.Length);
                        }
                        context.Response.Close();
                    }
                    else if (context.Request.HttpMethod.Equals(@"POST") && 
                        context.Request.Headers["X-WOPI-Override"].Equals("COBALT"))
                    {
                        Console.WriteLine("6666...");
                        //cobalt, for docx and pptx
                        var ms = new MemoryStream();
                        context.Request.InputStream.CopyTo(ms);
                        AtomFromByteArray atomRequest = new AtomFromByteArray(ms.ToArray());
                        RequestBatch requestBatch = new RequestBatch();

                        Object ctx;
                        ProtocolVersion protocolVersion;

                        requestBatch.DeserializeInputFromProtocol(atomRequest, out ctx, out protocolVersion);
                        editSession.ExecuteRequestBatch(requestBatch);

                        foreach (Request request in requestBatch.Requests)
                        {
                            if (request.GetType() == typeof(PutChangesRequest) && 
                                request.PartitionId == FilePartitionId.Content)
                            {
                               
                                editSession.Save();
                             
                            }
                        }
                        var response = requestBatch.SerializeOutputToProtocol(protocolVersion);

                        context.Response.Headers.Add("X-WOPI-CorellationID", context.Request.Headers["X-WOPI-CorrelationID"]);
                        context.Response.Headers.Add("request-id", context.Request.Headers["X-WOPI-CorrelationID"]);
                        context.Response.ContentType = @"application/octet-stream";
                        context.Response.ContentLength64 = response.Length;
                        response.CopyTo(context.Response.OutputStream);
                        context.Response.Close();
                    }
                    else if (context.Request.HttpMethod.Equals(@"POST") &&
                        (context.Request.Headers["X-WOPI-Override"].Equals("LOCK") ||
                        context.Request.Headers["X-WOPI-Override"].Equals("UNLOCK") ||
                        context.Request.Headers["X-WOPI-Override"].Equals("REFRESH_LOCK"))
                        )
                    {
                        Console.WriteLine("7777...");
                        //lock, for xlsx and pptx
                        context.Response.ContentLength64 = 0;
                        context.Response.ContentType = @"text/html";
                        context.Response.StatusCode = (int)HttpStatusCode.OK;
                        context.Response.Close();
                    }
                    else
                    {
                        Console.WriteLine(@"Invalid request parameters");
                        ErrorResponse(context, @"Invalid request cobalt parameter");
                    }

                    Console.WriteLine("ok...");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(@"process request exception:" + ex.Message);
                }
                m_listener.BeginGetContext(ProcessRequest, m_listener);
            }
            catch (Exception ex)
            {
                Console.WriteLine(@"get request context:" + ex.Message);
                return;
            }
        }

        IDFSHandle _handle = new GridFSHandle();
        private Stream GetFileById(string fileId)
        {
            return _handle.GetFile(fileId);
        }

        private string GetFileName(string fileId)
        {
            return _handle.GetFileName(fileId);
        }

        public void StreamToFile(Stream stream, string fileName)
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["FileSavePath"];
            path = Path.Combine(path, fileName);

            if(File.Exists(path))
            {
                File.Delete(path);
            }

            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            stream.Seek(0, SeekOrigin.Begin);

            // 把 byte[] 写入文件 
            using(FileStream fs = new FileStream(path, FileMode.Create))
            {
                using(BinaryWriter bw = new BinaryWriter(fs))
                {
                    bw.Write(bytes);
                }
            }
        }
    }
}
