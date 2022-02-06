using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

using Newtonsoft.Json;

namespace Infinium.Modules.Marketing.NewOrders
{

    public class CurrencyConverter
    {
        private class Currency
        {
            public int Cur_ID { get; set; }
            public DateTime Date { get; set; }
            public string Cur_Abbreviation { get; set; }
            public int Cur_Scale { get; set; }
            public string Cur_Name { get; set; }
            public decimal? Cur_OfficialRate { get; set; }
        }

        public static decimal NbrbDailyRates(DateTime date)
        {
            string url = $"https://www.nbrb.by/api/exrates/rates?periodicity=0&ondate={date:yyyy-MM-dd}";

            HttpWebResponse myHttpWebResponse = null;

            decimal eur = 0;
            decimal usd = 0;
            decimal rub = 0;

            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                var myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                myHttpWebRequest.UseDefaultCredentials = true;
                myHttpWebRequest.KeepAlive = false;
                myHttpWebRequest.AllowAutoRedirect = true;
                CookieContainer cookieContainer = new CookieContainer();
                myHttpWebRequest.CookieContainer = cookieContainer;
                myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
            }
            catch (NotSupportedException e)
            {

            }
            catch (ProtocolViolationException e)
            {

            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {

                }
            }
            catch (Exception e)
            {

            }
            finally
            {
            }

            List<Currency> list = new List<Currency>();
            if (myHttpWebResponse != null)
            {
                using (var reader = new StreamReader(myHttpWebResponse.GetResponseStream()))
                {
                    string objText = reader.ReadToEnd();
                    //list = new JavaScriptSerializer().Deserialize<List<Currency>>(objText);
                    list = JsonConvert.DeserializeObject<List<Currency>>(objText);
                }
            }

            if (list.Count > 0)
            {
                eur = (decimal)list.SingleOrDefault(p => p.Cur_Abbreviation == "EUR").Cur_OfficialRate;
                usd = (decimal)list.SingleOrDefault(p => p.Cur_Abbreviation == "USD").Cur_OfficialRate;
                rub = (decimal)list.SingleOrDefault(p => p.Cur_Abbreviation == "RUB").Cur_OfficialRate;
            }

            return eur;
        }
    }
}
