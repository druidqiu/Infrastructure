using System;
using System.IO;
using System.Text;
using System.Web;

namespace SYDQ.Infrastructure.Common
{
    public class ResponseHelper
    {
        public static void ResponseFile(string fileFullName, string fileDisplayName, string filePostfix)
        {
            byte[] bytes = ConvertFileToBytes(fileFullName);
            if (bytes != null && bytes.Length > 0)
            {
                ResponseWriteByteFile(bytes, fileDisplayName, filePostfix);
            }
        }

        public static void ResponseFileAndDelete(string fileFullName, string fileDisplayName, string filePostfix)
        {
            var bytes = ConvertFileToBytes(fileFullName);
            File.Delete(fileFullName);
            if (bytes != null && bytes.Length > 0)
            {
                ResponseWriteByteFile(bytes, fileDisplayName, filePostfix);
            }
        }

        private static void ResponseWriteByteFile(byte[] wordContent, string fileDisplayName, string filePostfix)
        {
            if (string.IsNullOrEmpty(fileDisplayName) || string.IsNullOrEmpty(filePostfix))
            {
                return;
            }

            var response = HttpContext.Current.Response;

            long dataToRead = wordContent.Length;
            const int length = 10000;
            var start = 0;

            response.Clear();
            response.ClearHeaders();
            response.Buffer = false;

            response.ContentType = "application/octet-stream";
            var urlEncode = HttpUtility.UrlEncode(fileDisplayName + filePostfix, Encoding.UTF8);
            if (urlEncode != null)
            {
                response.AddHeader("Content-Disposition", "attachment; filename=" + urlEncode.Replace("+", "%20"));
            }
            response.AddHeader("Content-Length", dataToRead.ToString());

            try
            {
                while (dataToRead > 0)
                {
                    if (response.IsClientConnected)
                    {
                        int reallength = dataToRead - length < 0 ? (int)dataToRead : length;
                        response.OutputStream.Write(wordContent, start, reallength);
                        response.Flush();

                        start += reallength;
                        dataToRead = dataToRead - reallength;
                    }
                    else
                    {
                        dataToRead = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                response.Clear();
                response.Write(ex.Message);
                response.End();
            }
            finally
            {
                response.Clear();
                response.End();
            }
        }

        private static byte[] ConvertFileToBytes(string fileFullName)
        {
            byte[] bytContent = null;
            FileStream stream = null;
            BinaryReader reader = null;
            try
            {
                stream = new FileStream(fileFullName, FileMode.Open);

                reader = new BinaryReader((Stream)stream);
                bytContent = reader.ReadBytes((Int32)stream.Length);
            }
            catch (Exception ex)
            {
                //TODO:
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
                if (reader != null)
                {
                    reader.Close();
                }
            }

            return bytContent;
        }
    }
}
