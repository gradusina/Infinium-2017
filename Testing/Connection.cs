using System;
using System.Collections;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Infinium;

namespace Testing
{
    public class SplashWindow
    {
        public static bool bSmallCreated = false;

        public static void CreateSplash()
        {
            SplashForm SplashForm = new SplashForm();

            SplashForm.ShowDialog();

            SplashForm.Dispose();
            SplashForm = null;
            GC.Collect();
        }

        public static void CreateSmallSplash(string Message)
        {
            SmallWaitForm SmallWaitForm = new SmallWaitForm(Message);

            bSmallCreated = true;

            SmallWaitForm.ShowDialog();

            SmallWaitForm.Dispose();
            bSmallCreated = false;
            SmallWaitForm = null;
            GC.Collect();
        }

        public static void CreateSmallSplash(ref Form TopForm, string Message)
        {
            SmallWaitForm SmallWaitForm = new SmallWaitForm(Message);

            bSmallCreated = true;

            SmallWaitForm.ShowDialog();

            SmallWaitForm.Dispose();
            bSmallCreated = false;
            SmallWaitForm = null;
            GC.Collect();
        }

        public static void CreateCoverSplash(int Top, int Left, int Height, int Width)
        {
            CoverWaitForm CoverWaitForm = new CoverWaitForm(false)
            {
                Width = Width,
                Height = Height,
                Top = Top,
                Left = Left
            };
            CoverWaitForm.ShowDialog();

            CoverWaitForm.Dispose();
            bSmallCreated = false;
            CoverWaitForm = null;
            GC.Collect();
        }

        public static void CreateCoverSplash(bool bSmall, int Top, int Left, int Height, int Width)
        {
            CoverWaitForm CoverWaitForm = new CoverWaitForm(bSmall)
            {
                Width = Width,
                Height = Height,
                Top = Top,
                Left = Left
            };
            CoverWaitForm.ShowDialog();

            CoverWaitForm.Dispose();
            bSmallCreated = false;
            CoverWaitForm = null;
            GC.Collect();
        }

        public static void CreateCoverSplash(int Top, int Left, int Height, int Width, Color BackColor)
        {
            CoverWaitForm CoverWaitForm = new CoverWaitForm(false, BackColor)
            {
                Width = Width,
                Height = Height,
                Top = Top,
                Left = Left
            };
            CoverWaitForm.ShowDialog();

            CoverWaitForm.Dispose();
            bSmallCreated = false;
            CoverWaitForm = null;
            GC.Collect();
        }
    }
    public class DatabaseConfigsManager
    {
        public static bool Animation;

        private string sConnectionString = null;

        public DatabaseConfigsManager()
        {

        }

        public void ReadAnimationFlag(String FileName)
        {
            using (System.IO.StreamReader sr = new System.IO.StreamReader(FileName))
            {
                Animation = Convert.ToBoolean(sr.ReadToEnd());
            }
        }

        public string ReadConfig(String FileName)
        {
            using (System.IO.StreamReader sr = new System.IO.StreamReader(FileName))
            {
                sConnectionString = sr.ReadToEnd();
            }

            return sConnectionString;
        }

        public string ReadConfig(String FileName, int BytesToRead, int StartByte)
        {
            using (System.IO.StreamReader sr = new System.IO.StreamReader(FileName))
            {
                string s = sr.ReadToEnd();

                sConnectionString = s.Substring(StartByte, BytesToRead);
            }

            return sConnectionString;
        }
    }

    public class Connection
    {
        private string sConnectionString = null;

        public Connection()
        {

        }

        private string ReadConnectionString(String FileName)
        {
            using (System.IO.StreamReader sr = new System.IO.StreamReader(FileName))
            {
                sConnectionString = sr.ReadToEnd();
            }

            return sConnectionString;
        }

        private string EncryptString(string inputString, int dwKeySize, string xmlString)
        {
            RSACryptoServiceProvider rsaCryptoServiceProvider = new RSACryptoServiceProvider(dwKeySize);
            rsaCryptoServiceProvider.FromXmlString(xmlString);
            int keySize = dwKeySize / 8;
            byte[] bytes = Encoding.UTF32.GetBytes(inputString);

            int maxLength = keySize - 42;
            int dataLength = bytes.Length;
            int iterations = dataLength / maxLength;
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i <= iterations; i++)
            {
                byte[] tempBytes = new byte[(dataLength - maxLength * i > maxLength) ? maxLength : dataLength - maxLength * i];
                Buffer.BlockCopy(bytes, maxLength * i, tempBytes, 0, tempBytes.Length);
                byte[] encryptedBytes = rsaCryptoServiceProvider.Encrypt(tempBytes, true);

                Array.Reverse(encryptedBytes);

                stringBuilder.Append(Convert.ToBase64String(encryptedBytes));
            }
            return stringBuilder.ToString();
        }

        private string DecryptString(string inputString, int dwKeySize, string xmlString)
        {
            RSACryptoServiceProvider rsaCryptoServiceProvider = new RSACryptoServiceProvider(dwKeySize);
            rsaCryptoServiceProvider.FromXmlString(xmlString);
            int base64BlockSize = ((dwKeySize / 8) % 3 != 0) ? (((dwKeySize / 8) / 3) * 4) + 4 : ((dwKeySize / 8) / 3) * 4;
            int iterations = inputString.Length / base64BlockSize;
            ArrayList arrayList = new ArrayList();
            for (int i = 0; i < iterations; i++)
            {
                byte[] encryptedBytes = Convert.FromBase64String(inputString.Substring(base64BlockSize * i, base64BlockSize));

                Array.Reverse(encryptedBytes);
                arrayList.AddRange(rsaCryptoServiceProvider.Decrypt(encryptedBytes, true));
            }
            return Encoding.UTF32.GetString(arrayList.ToArray(Type.GetType("System.Byte")) as byte[]);
        }

        public string GetConnectionString(string FileName)
        {
            try
            {
                string EncryptedConnectionString = ReadConnectionString(FileName);
                int Index = EncryptedConnectionString.IndexOf("Password");
                string Password = EncryptedConnectionString.Substring(Index + 9, EncryptedConnectionString.Length - Index - 9);
                string privateKey = "<RSAKeyValue><Modulus>pqJ+ilhHSM5N3XGPCAYdrpFjHlvcQQoaNPiTvUHut5dIwx40olIKRjvequY8WeGb</Modulus><Exponent>AQAB</Exponent><P>2UlRsj0mJoeGxD4AtwOEn02oCEjlYifD</P><Q>xFLlKWJuoE3mb88iI8v24/7Qt1Wvc5lJ</Q><DP>s8j4sfPqpyKoHaP3z3Y3u9/zUreOJIMl</DP><DQ>YhIOy8eR/54qeLv+D+e5o1cNKCgzhwmR</DQ><InverseQ>kU9Phb5ynWsB6ZFQoAnUPAzmirRIlqDR</InverseQ><D>J1Q64ZQsXvayUg23YIFxB/6wkj3EImWroC3gjjCvYa+fjojM1XXvE/tE5t8mnnzB</D></RSAKeyValue>";
                sConnectionString = EncryptedConnectionString.Replace(Password, DecryptString(Password, 384, privateKey));
            }
            catch (ArgumentNullException)
            {

            }
            return sConnectionString;
        }

        public string DecryptStringConnectionString(string EncryptedConnectionString)
        {
            try
            {
                int Index = EncryptedConnectionString.IndexOf("Password");
                string Password = EncryptedConnectionString.Substring(Index + 9, EncryptedConnectionString.Length - Index - 9);
                string privateKey = "<RSAKeyValue><Modulus>pqJ+ilhHSM5N3XGPCAYdrpFjHlvcQQoaNPiTvUHut5dIwx40olIKRjvequY8WeGb</Modulus><Exponent>AQAB</Exponent><P>2UlRsj0mJoeGxD4AtwOEn02oCEjlYifD</P><Q>xFLlKWJuoE3mb88iI8v24/7Qt1Wvc5lJ</Q><DP>s8j4sfPqpyKoHaP3z3Y3u9/zUreOJIMl</DP><DQ>YhIOy8eR/54qeLv+D+e5o1cNKCgzhwmR</DQ><InverseQ>kU9Phb5ynWsB6ZFQoAnUPAzmirRIlqDR</InverseQ><D>J1Q64ZQsXvayUg23YIFxB/6wkj3EImWroC3gjjCvYa+fjojM1XXvE/tE5t8mnnzB</D></RSAKeyValue>";
                sConnectionString = EncryptedConnectionString.Replace(Password, DecryptString(Password, 384, privateKey));
            }
            catch (ArgumentNullException)
            {

            }
            return sConnectionString;
        }
    }





    public struct ConnectionStrings
    {
        public static string CatalogConnectionString;
        public static string LightConnectionString;
        public static string MarketingOrdersConnectionString;
        public static string MarketingReferenceConnectionString;
        public static string StorageConnectionString;
        public static string UsersConnectionString;
        public static string ZOVOrdersConnectionString;
        public static string ZOVReferenceConnectionString;
    }
}
