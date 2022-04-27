using Inventariz.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Report;
using Newtonsoft.Json;
using System.Web.UI.DataVisualization.Charting;

namespace Inventariz.Controllers
{
    public class HIM
    {
        public DateTime? dat;
        public int? rezer;
        public double? densreal;
        public double? dens20;
        public double? tempreal;
        public double? water;
        public double? saltmg;
        public double? mechan;
    }
    public class TankInfoType
    {
        public int tankid;
        public DateTime date;
        public int filial;
        public double level;
        public double temp;
        public int type;
    }



    public class HomeController : Controller
    {
        public asopnEntities db = new asopnEntities();
        public ChemLabEntities db2 = new ChemLabEntities();
        trl_tankview vi = new trl_tankview();

        public ActionResult Index()
        {
            string Hour = DateTime.Now.Hour.ToString();
            string Dat = DateTime.Now.ToShortDateString().ToString();
            string DatHour = Hour + ":00" + "; " + Dat;
            ViewBag.DatHour = DatHour;

            //----Получаем суммарный объем нефти в РП Мозырь-------------------

            List<TankInv> TableInv = new List<TankInv>();
            List<TankInv> TableInvNov = new List<TankInv>();
            DateTime las = db.TankInv.Max(d => d.Data);
            TableInv = db.TankInv.Where(j => j.Filial == 1 & j.Data == las).ToList();
            TableInvNov = db.TankInv.Where(j => j.Filial == 2 & j.Data == las).ToList(); ;


            double VNeftItog1 = 0;
            double VNeftItog2 = 0;
            
            foreach (var t in TableInv)
            {                
                VNeftItog1 = VNeftItog1 + Convert.ToDouble(t.VNeft);                
            }
            //---------------------------------------------
            ViewBag.VNeftItog1 = VNeftItog1;

            foreach (var t in TableInvNov)
            {
                VNeftItog2 = VNeftItog2 + Convert.ToDouble(t.VNeft);
            }
            //---------------------------------------------
            ViewBag.VNeftItog2 = VNeftItog2;
            ViewBag.SumNeft = VNeftItog1 + VNeftItog2;

            //-----------------------------------------------------------------

            return View();
        }

        public ActionResult Calc()
        {

            SelectList podrazdD = new SelectList(db.filials, "id", "name");
            ViewBag.podr = podrazdD;

            SelectList RezMoz = new SelectList(db.trl_tankview, "TankID", "TankName");
            ViewBag.rezer = RezMoz;


            return View();
        }

        public ActionResult GradTables()
        {
            SelectList podrazdD = new SelectList(db.filials, "id", "name");
            ViewBag.podr = podrazdD;

            SelectList RezMoz = new SelectList(db.trl_tankview, "TankID", "TankName");
            ViewBag.rezer = RezMoz;

            return View();
        }

        public ActionResult AktInvent()
        {
            SelectList podrazdD1 = new SelectList(db.filials, "id", "name");
            ViewBag.podr1 = podrazdD1;

            List<TankInv> DatINV = new List<TankInv>();
            DatINV = db.TankInv.Where(j => j.Filial == 1).OrderByDescending(h => h.Data).ToList();
            List<TankInv> DatInv = new List<TankInv>();

            var maxDatG = db.TankInv.Where(k => k.Filial == 1).Max(y => y.Data);
            DatInv = db.TankInv.Where(j => j.Filial == 1 & j.Data == maxDatG).OrderBy(h => h.Rezer).ToList();

            SelectList inventGomel = new SelectList(db.filials, "id", "name");
            ViewBag.datinv = DatInv;

            List<DateTime?> datSelect = new List<DateTime?>();

            foreach (var item in DatINV)
            {
                datSelect.Add(item.Data);
            }
            List<DateTime?> datSelectDiST = new List<DateTime?>();
            datSelectDiST = datSelect.Distinct().ToList();
            ViewBag.SelectDist = datSelectDiST;

            return View();
        }

        public ActionResult NormInformation()
        {

            return View();
        }

        //----------------Подписанты--------------------------//
        public ActionResult Kom()
        {
            List<Podpisanty> podpisanty = new List<Podpisanty>();
            podpisanty = db.Podpisanty.OrderBy(h => h.IDFilial).ToList();

            SelectList komis = new SelectList(db.filials, "id", "name");
            ViewBag.komis = komis;

            return View(podpisanty);
        }

        //----------Добавление члена комиссии-----------------------//

        public ActionResult AddKomissiya()
        {
            SelectList komis = new SelectList(db.filials, "id", "name");
            ViewBag.komis = komis;
            return PartialView();
        }
        //--------------------------//
        //-----Сохранение добавления члена комиссии-----------//
        [HttpPost]
        public ActionResult KomissiyaSave(string location, string Nazn, string Doljnost, string FIO)
        {
            try
            {
                Podpisanty kom = new Podpisanty();
                kom.IDFilial = Convert.ToInt32(location);
                kom.Nazn = Nazn.Trim();
                kom.Doljnost = Doljnost.Trim();
                kom.FIO = FIO.Trim();
                db.Podpisanty.Add(kom);

                db.SaveChanges();

                ViewBag.Message = "Член комиссии успешно добавлендобавлен!";
            }
            catch (Exception ex)
            {


                ViewBag.Message = ex.ToString();
            }

            return PartialView();
        }
        //----------------------------------//
        // удаление подписанта//

        public ActionResult DeletePodpis(int ID)
        {
            Podpisanty podp = new Podpisanty();
            podp = db.Podpisanty.FirstOrDefault(a => a.ID == ID);
            return PartialView(podp);
        }
        //-----------------------------//

        // Подтверждение удаления подписанта//
        public ActionResult DeletePodpisOK(int ID)
        {
            try
            {

                Podpisanty podpis = new Podpisanty();
                podpis = db.Podpisanty.FirstOrDefault(a => a.ID == ID);
                db.Podpisanty.Remove(podpis);
                db.SaveChanges();

                ViewBag.Message = "Подписант успешно удален!";
            }
            catch
            {
                ViewBag.Message = "Ошибка удаления";
            }

            return PartialView();
        }
        // Редактирование члена комиссии//

        public ActionResult KomissiyaEdit(int ID)
        {
            Podpisanty pod = new Podpisanty();
            pod = db.Podpisanty.FirstOrDefault(a => a.ID == ID);
            return PartialView(pod);
        }
        //-------------------------------//

        //Сохранение редактирования члена комиссии------------//
        [HttpPost]
        public ActionResult KomissiyaEditSave(int ID, string Nazn, string Doljnost, string FIO)
        {
            try
            {
                Podpisanty pod = new Podpisanty();
                pod = db.Podpisanty.FirstOrDefault(s => s.ID == ID);

                pod.Nazn = Nazn.Trim();
                pod.Doljnost = Doljnost.Trim();
                pod.FIO = FIO.Trim();
                db.SaveChanges();

                ViewBag.Message = "Подписант успешно изменен";

            }
            catch (Exception e)
            {
                ViewBag.Message = "Ошибка. Текст ошибки: " + e.ToString();
            }

            return PartialView();
        }
        //-----------------------------//

        public ActionResult Reference()
        {

            return PartialView();
        }
        //------------------Список резервуаров в зависимости от филиала--------------------
        public ActionResult GetRezerv(int ID)
        {
            //List<string> Rezerv = new List<string>();
            ////Rezerv = db.trl_tankview.Where(b => b.name.Equals(ID)).Select(x => x.name).Distinct().ToList();
            //Rezerv = db.trl_tankview.Where(b => b.FilialID == ID).Select(f => f.TankName).ToList();

            List<trl_tankview> Rezerv = new List<trl_tankview>();
            Rezerv = db.trl_tankview.Where(h => h.FilialID == ID).ToList();

            return PartialView(Rezerv);
        }

        //-----------------Получение градуировочной таблицы в зависимоти от резервуара-----------------------------------
        public ActionResult GetGradTable(int rezer)
        {
            List<calibration> GetTab = new List<calibration>();
            GetTab = db.calibration.Where(j => j.tankid == rezer).OrderBy(h => h.id).ToList();

            ViewBag.GetTab = GetTab;

            return PartialView(GetTab);
        }

        //----------------------------------//
        // удаление градуировочной таблицы//

        public ActionResult DeleteGrad(int ID)
        {
            calibration calibr = new calibration();
            calibr = db.calibration.FirstOrDefault(a => a.id == ID);
            return PartialView(calibr);
        }
        //-----------------------------//

        // Подтверждение удаления градуировочной таблицы//
        public ActionResult DeleteGradOK(int ID)
        {
            try
            {

                calibration calibr = new calibration();
                calibr = db.calibration.FirstOrDefault(a => a.id == ID);
                db.calibration.Remove(calibr);
                db.SaveChanges();

                ViewBag.Message = "Строка успешно удалена!";
            }
            catch
            {
                ViewBag.Message = "Ошибка удаления";
            }

            return PartialView();
        }
        // Редактирование градуировочной таблицы//

        public ActionResult GradEdit(int ID)
        {
            calibration calibr = new calibration();
            calibr = db.calibration.FirstOrDefault(a => a.id == ID);
            return PartialView(calibr);
        }
        //-------------------------------//

        //Сохранение редактирования градуировочной таблицы------------//
        [HttpPost]
        public ActionResult GradEditSave(int ID, int urov, string V)
        {
            try
            {
                calibration calibr = new calibration();
                calibr = db.calibration.FirstOrDefault(s => s.id == ID);

                calibr.oillevel = urov;
                calibr.oilvolume = Convert.ToDouble(V);
                db.SaveChanges();

                ViewBag.Message = "Строка успешно изменена";

            }
            catch (Exception e)
            {
                ViewBag.Message = "Ошибка. Текст ошибки: " + e.ToString();
            }

            return PartialView();
        }
        //-----------------------------//
        //----------Добавление строки в градуировочную таблицу-----------------------//

        public ActionResult AddGradTable()
        {
            SelectList calibr = new SelectList(db.calibration, "id", "tankid");
            ViewBag.calibr = calibr;

            return PartialView();
        }
        //--------------------------//
        //-----Сохранение добавления строки в градуировочную таблицу-----------//
        [HttpPost]
        public ActionResult GradTableSave(int IDRez, int urov, string V)
        {
            try
            {
                calibration cal = new calibration();
                cal.tankid = IDRez;
                cal.oillevel = urov;
                cal.oilvolume = Convert.ToDouble(V);
                db.calibration.Add(cal);

                db.SaveChanges();

                ViewBag.Message = "Строка успешно добавлендобавлена!";
            }
            catch (Exception ex)
            {


                ViewBag.Message = ex.ToString();
            }

            return PartialView();
        }
        //----------------------------------//

        //---------Получение даты инвентаризации в зависимости от филиала-----------------------------
        public ActionResult GetDatInv(int ID)
        {

            List<TankInv> DatInv = new List<TankInv>();
            DatInv = db.TankInv.Where(h => h.Filial == ID).ToList();

            SelectList podrazdD1 = new SelectList(db.filials, "id", "name");
            ViewBag.podr1 = podrazdD1;

            List<TankInv> DatINV = new List<TankInv>();
            DatINV = db.TankInv.Where(j => j.Filial == ID).OrderByDescending(h => h.Data).ToList();
            List<TankInv> DatINVLast = new List<TankInv>();

            var maxDatG = db.TankInv.Where(k => k.Filial == 1).Max(y => y.Data);
            DatInv = db.TankInv.Where(j => j.Filial == ID & j.Data == maxDatG).OrderBy(h => h.Rezer).ToList();

            SelectList inventGomel = new SelectList(db.filials, "id", "name");
            ViewBag.datinv = DatInv;

            List<DateTime> datSelect = new List<DateTime>();

            foreach (var item in DatINV)
            {
                datSelect.Add(item.Data);
            }
            List<DateTime> datSelectDiST = new List<DateTime>();
            datSelectDiST = datSelect.Distinct().ToList();
            ViewBag.SelDat = datSelectDiST;

            return PartialView();
        }
        //--------------------------------------------------------------------------------------------
        //-------------------Вывод таблицы акта ивентаризации-----------------------------------------
        public ActionResult GetTableInv(DateTime datinv, int filial1)
        {
            //-----------------------------------------------------------------------------------------------------------
            List<TankInv> TableInv = new List<TankInv>();
            TableInv = db.TankInv.Where(j => j.Filial == filial1 & j.Data == datinv).OrderBy(h => h.type).ThenBy(h => h.Rezer).ToList();

            //список тиров резервуаров
            //--------------------------------------------------
            List<trl_tank> trl = new List<trl_tank>();
            trl = db.trl_tank.Where(p => p.FilialID == filial1).OrderBy(p => p.TypeID).ThenBy(p => p.TankID).ToList();

            List<int> typ = new List<int>();
            foreach (var t in trl)
            {
                {
                    if (t.TypeID != 0)
                        typ.Add(t.TypeID);
                }
            }

            List<int> typDyst = new List<int>();

            typDyst = typ.Distinct().ToList();
            ViewBag.typ = typDyst;

            //--------------------------------------------
            int kol = 0;
            double VNeftItog1 = 0;
            double PSred1 = 0;
            double MassaBalItog1 = 0;
            double BalProcSred1 = 0;
            double BalTonnItog1 = 0;
            double MassaNettoItog1 = 0;
            double MassaBalMinItog1 = 0;
            double MassaNettoMinItog1 = 0;

            foreach (var t in TableInv)
            {
                kol++;
                VNeftItog1 = VNeftItog1 + Convert.ToDouble(t.VNeft);
                PSred1 = Convert.ToDouble(t.P) + PSred1;
                MassaBalItog1 = MassaBalItog1 + Convert.ToDouble(t.MassaBrutto);
                BalProcSred1 = BalProcSred1 + Convert.ToDouble(t.BalProc);
                BalTonnItog1 = BalTonnItog1 + Convert.ToDouble(t.BalTonn);
                MassaNettoItog1 = MassaNettoItog1 + Convert.ToDouble(t.MassaNetto);
                MassaBalMinItog1 = MassaBalMinItog1 + Convert.ToDouble(t.MBalMin);
                MassaNettoMinItog1 = MassaNettoMinItog1 + Convert.ToDouble(t.MNettoMin);
            }
            //---------------------------------------------
            ViewBag.VNeftItog1 = VNeftItog1;
            ViewBag.PSred1 = Math.Round(PSred1 / kol, 1);
            ViewBag.MassaBalItog1 = Math.Round(MassaBalItog1,1);
            ViewBag.BalProcSred1 = Math.Round(BalProcSred1 / kol, 1);
            ViewBag.BalTonnItog1 = Math.Round(BalTonnItog1, 1);
            ViewBag.MassaNettoItog1 = Math.Round(MassaNettoItog1, 1);
            ViewBag.MassaBalMinItog1 = Math.Round(MassaBalMinItog1, 1);
            ViewBag.MassaNettoMinItog1 = Math.Round(MassaNettoMinItog1, 1);

            ViewBag.TableInv = TableInv;

            return PartialView(TableInv);
            //------------------------------------------------------------------------------------------------------------
        }
        //--------------------------------------------------------------------------------------------               

        //----------Получить химанализ в зависимости от резервуара-------------------------------
        public ActionResult GetHim(int filial, int rezer)
        {
            //string connectionString = @"Persist Security Info = True; User ID = inventor; Initial Catalog = asopn; Data Source = askid; Password = inventor";
            //using (SqlConnection conn = new SqlConnection(connectionString))
            //{ 
            //        string qry = $"SELECT TOP 1 p.sampledt, p.samplelocationstr, r.densreal, r.dens20,r.tempreal, r.water, r.saltmg, r.mechan FROM dbo.Protocol p" +
            //               $"JOIN dbo.ProtocolResult r ON r.protocolid=p.id" +
            //               $"WHERE p.samplelocationid=15 ORDER BY p.sampledt desc";
            //        SqlCommand command = new SqlCommand(qry, conn);

            //    try
            //    {
            //        conn.Open();
            //        SqlDataReader reader = command.ExecuteReader();
            //        var list1 = new List<him>();
            //        while (reader.Read())
            //        {
            //            var d = new him();
            //            d.dates = reader.GetDateTime(1);
            //            d.P = reader.GetDouble(2);
            //            d.temp = reader.GetDouble(3);
            //            d.P20 = reader.GetDouble(4);
            //            d.water = reader.GetDouble(5);
            //            d.saltmg = reader.GetDouble(6);
            //            d.mechan = reader.GetDouble(7);
            //            list1.Add(d);
            //        }
            //        reader.Close();
            //        return PartialView(list1);
            //    }
            //    catch (Exception ex)
            //    {

            //        return PartialView(ex);
            //    }
            //}
            HIM H = new HIM();
            if (filial == 1)
            {
                LastTanksResultMoz MOZ = new LastTanksResultMoz();
                MOZ = db2.LastTanksResultMoz.FirstOrDefault(f => f.tankid == rezer);

                if (MOZ == null)
                {
                    H.dat = null;
                    H.densreal = null;
                    H.dens20 = null;
                    H.tempreal = null;
                    H.water = null;
                    H.saltmg = null;
                    H.mechan = null;

                }
                else
                {

                    H.dat = MOZ.endsampledt;
                    H.densreal = MOZ.densreal;
                    H.dens20 = MOZ.dens20;
                    H.tempreal = Convert.ToDouble(MOZ.tempreal);
                    H.water = MOZ.water;
                    H.saltmg = MOZ.saltmg;
                    H.mechan = MOZ.mechan;
                }
            }
            if (filial == 2)
            {
                LastTanksResultPol MOZ = new LastTanksResultPol();
                MOZ = db2.LastTanksResultPol.FirstOrDefault(j => j.tankid == rezer);
                if (MOZ == null)
                {
                    H.dat = null;
                    H.densreal = null;
                    H.dens20 = null;
                    H.tempreal = null;
                    H.water = null;
                    H.saltmg = null;
                    H.mechan = null;

                }
                else
                {
                    H.dat = MOZ.sampledt;
                    H.densreal = MOZ.densreal;
                    H.dens20 = MOZ.dens20;
                    H.tempreal = Convert.ToDouble(MOZ.tempreal);
                    H.water = MOZ.water;
                    H.saltmg = MOZ.saltmg;
                    H.mechan = MOZ.mechan;
                }

            }
            ViewBag.dat = H.dat;
            ViewBag.dens = H.densreal;
            ViewBag.temp = H.tempreal;
            ViewBag.dens20 = H.dens20;

            return PartialView();
        }

        //----------Калькулятор разность объемов-------------------------------
        public ActionResult GetRazn(int filial1, int rezer1, string Unach, string Ukon, string UnachH2O, string UkonH2O, string Pnach, string Pkon, string Tnach, string Tkon, string Bal, string Kst, string Ksr)
        {
            HIM H = new HIM();
            double V1;
            double V1R;
            double V1RN;
            double V2;
            double V2R;
            double V2RN;
            double RV;
            double RM;
            double Unachmin;
            double Unachmax;
            double Vnachmin;
            double Vnachmax;
            double Upercentnach;
            double Mnach;
            double RMB;

            //Для начальной подтоварной воды
            double UnachminH2O;
            double UnachmaxH2O;
            double VnachminH2O;
            double VnachmaxH2O;
            double V1H2O;
            double V1RH2O;
            double UpercentnachH2O;
            //--------------------------------

            double Ukonmin;
            double Ukonmax;
            double Vkonmin;
            double Vkonmax;
            double Upercentkon;
            double Mkon;

            //Для подтоварной воды конечной
            double UkonminH2O;
            double UkonmaxH2O;
            double VkonminH2O;
            double VkonmaxH2O;
            double V2H2O;
            double V2RH2O;
            double UpercentkonH2O;
            //--------------------------------

            List<calibration> calibrationList = new List<calibration>();
            calibrationList = db.calibration.Where(f => f.filialid == filial1).OrderBy(f => f.tankid).ThenBy(f => f.oillevel).ToList();

            if (filial1 == 1)
            {
                LastTanksResultMoz MOZ = new LastTanksResultMoz();
                MOZ = db2.LastTanksResultMoz.FirstOrDefault(f => f.tankid == rezer1);
                if (MOZ == null)
                {
                    H.dat = null;
                    H.densreal = null;
                    H.dens20 = null;
                    H.tempreal = null;
                    H.water = null;
                    H.saltmg = null;
                    H.mechan = null;
                }
                else
                {
                    H.dat = MOZ.endsampledt;
                    H.densreal = MOZ.densreal;
                    H.dens20 = MOZ.dens20;
                    H.tempreal = Convert.ToDouble(MOZ.tempreal);
                    H.water = MOZ.water;
                    H.saltmg = MOZ.saltmg;
                    H.mechan = MOZ.mechan;
                }
            }

            if (filial1 == 2)
            {
                LastTanksResultPol MOZ = new LastTanksResultPol();
                MOZ = db2.LastTanksResultPol.FirstOrDefault(j => j.tankid == rezer1);
                if (MOZ == null)
                {
                    H.dat = null;
                    H.densreal = null;
                    H.dens20 = null;
                    H.tempreal = null;
                    H.water = null;
                    H.saltmg = null;
                    H.mechan = null;
                }
                else
                {
                    H.dat = MOZ.sampledt;
                    H.densreal = MOZ.densreal;
                    H.dens20 = MOZ.dens20;
                    H.tempreal = Convert.ToDouble(MOZ.tempreal);
                    H.water = MOZ.water;
                    H.saltmg = MOZ.saltmg;
                    H.mechan = MOZ.mechan;
                }
            }

            if (calibrationList.Where(h => h.tankid == rezer1).Count() == 0)
            {
                ViewBag.V1 = "данных нет!";
                ViewBag.V2 = "даных нет!";
                ViewBag.Mnach = "данных нет!";
                ViewBag.Mkon = "данных нет!";
                ViewBag.RV = "данных нет!";
                ViewBag.RM = "данных нет!";
            }
            else if (calibrationList.LastOrDefault(g => g.tankid == rezer1).oillevel < Convert.ToDouble(Unach) | calibrationList.LastOrDefault(g => g.tankid == rezer1).oillevel < Convert.ToDouble(Ukon))
            {
                ViewBag.V1 = "Большое значение!";
                ViewBag.V2 = "Большое значение!";
                ViewBag.Mnach = "Большое значение!";
                ViewBag.Mkon = "Большое значение!";
                ViewBag.RV = "Большое значение!";
                ViewBag.RM = "Большое значение!";
            }
            else
            {

                //-------Рассчет объема нефти начального---------------------------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Unach)) == null)
                {
                    Unachmin = 0;
                }
                else
                {
                    Unachmin = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Unach)).oillevel;
                }
                Unachmax = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(Unach)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Unach)) == null)
                {
                    Vnachmin = 0;
                }
                else
                {
                    Vnachmin = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Unach)).oilvolume;
                }
                Vnachmax = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(Unach)).oilvolume;
                Upercentnach = (Convert.ToDouble(Unach) - Unachmin) / (Unachmax - Unachmin);
                V1 = Vnachmin + (Vnachmax - Vnachmin) * Upercentnach; //по градуировочной таблице
                V1R = V1 * (1 + (2 * Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tnach) - 20)); //рассчитано по формуле

                //---------Рассчет объема подтоварной воды начальной-------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UnachH2O)) == null)
                {
                    UnachminH2O = 0;
                }
                else
                {
                    UnachminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UnachH2O)).oillevel;
                }
                UnachmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(UnachH2O)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UnachH2O)) == null | Convert.ToDouble(UnachH2O) == 0)
                {
                    VnachminH2O = 0;
                }
                else
                {
                    VnachminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UnachH2O)).oilvolume;
                }
                VnachmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(UnachH2O)).oilvolume;
                UpercentnachH2O = (Convert.ToDouble(UnachH2O) - UnachminH2O) / (UnachmaxH2O - UnachminH2O);
                V1H2O = VnachminH2O + (VnachmaxH2O - VnachminH2O) * UpercentnachH2O;
                V1RH2O = V1H2O * (1 + (Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tnach) - 20));

                V1RN = (Math.Round(V1, 3) - Math.Round(V1H2O, 3)) * (1 + (2 * Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tnach) - 20)); // Рассчет конечного объема чистой нефти без подтоварной воды

                //---------Рассчет объема нефти конечного----------------------------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Ukon)) == null)
                {
                    Ukonmin = 0;
                }
                else
                {
                    Ukonmin = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Ukon)).oillevel;
                }
                Ukonmax = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(Ukon)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Ukon)) == null)
                {
                    Vkonmin = 0;
                }
                else
                {
                    Vkonmin = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(Ukon)).oilvolume;
                }
                Vkonmax = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(Ukon)).oilvolume;
                Upercentkon = (Convert.ToDouble(Ukon) - Ukonmin) / (Ukonmax - Ukonmin);
                V2 = Vkonmin + (Vkonmax - Vkonmin) * Upercentkon;
                V2R = V2 * (1 + (Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tkon) - 20));

                //------------Рассчет объема подтоварной воды конечной-----------------------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UkonH2O)) == null)
                {
                    UkonminH2O = 0;
                }
                else
                {
                    UkonminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UkonH2O)).oillevel;
                }
                UkonmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(UkonH2O)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UkonH2O)) == null | Convert.ToDouble(UkonH2O) == 0)
                {
                    VkonminH2O = 0;
                }
                else
                {
                    VkonminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer1 & g.oillevel <= Convert.ToDouble(UkonH2O)).oilvolume;
                }
                VkonmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer1 & g.oillevel > Convert.ToDouble(UkonH2O)).oilvolume;
                UpercentkonH2O = (Convert.ToDouble(UkonH2O) - UkonminH2O) / (UkonmaxH2O - UkonminH2O);
                V2H2O = VkonminH2O + (VkonmaxH2O - VkonminH2O) * UpercentkonH2O;
                V2RH2O = V2H2O * (1 + (Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tkon) - 20));

                V2RN = (Math.Round(V2, 3) - Math.Round(V2H2O, 3)) * (1 + (2 * Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tkon) - 20)); //Расчет конечного объема читой нефти без подтоварной воды
                //--------------------------------------------------------------------------------------------------------------------
                Mnach = Convert.ToDouble(Pnach) * Math.Round(V1RN, 3) / 1000;
                Mkon = Convert.ToDouble(Pkon) * Math.Round(V2RN, 3) / 1000;
                RV = Math.Round(V2RN, 3) - Math.Round(V1RN, 3);
                RM = Mkon - Mnach;
                RMB = RM - RM * Convert.ToDouble(Bal) / 100;


                ViewBag.V1 = Math.Round(V1RN, 3);
                ViewBag.V2 = Math.Round(V2RN, 3);
                ViewBag.Mnach = Math.Round(Mnach, 3);
                ViewBag.Mkon = Math.Round(Mkon, 3);
                ViewBag.RV = Math.Round(RV, 0);
                ViewBag.RM = Math.Round(RM, 0);
                ViewBag.RMB = Math.Round(RMB, 0);
            }
            return PartialView();
        }

        public ActionResult GetRezerv1(int ID)
        {
            List<trl_tankview> Rezerv1 = new List<trl_tankview>();
            Rezerv1 = db.trl_tankview.Where(h => h.FilialID == ID).ToList();

            return PartialView(Rezerv1);
        }

        public ActionResult GetRezerv2(int ID)
        {
            List<trl_tankview> Rezerv2 = new List<trl_tankview>();
            Rezerv2 = db.trl_tankview.Where(h => h.FilialID == ID).ToList();

            return PartialView(Rezerv2);
        }

        //----------Калькулятор разность объемов по массе-------------------------------
        public ActionResult GetRazn1(int filial2, int rezer2, string Unach1, string RM1, string UnachH2O1, string UkonH2O1, string Pnach1, string Pkon1, string Tnach1, string Tkon1, string Bal1, string Kst1, string Ksr1)
        {
            HIM H = new HIM();
            double V1;
            double V1R;
            double V1RN;
            double V2;
            double V2R;
            double V2RN;
            double RV;
            double RM;
            double Unachmin;
            double Unachmax;
            double Vnachmin;
            double Vnachmax;
            double Upercentnach;
            double Mnach1;
            double RMB;

            //Для начальной подтоварной воды
            double UnachminH2O;
            double UnachmaxH2O;
            double VnachminH2O;
            double VnachmaxH2O;
            double V1H2O;
            double V1RH2O;
            double UpercentnachH2O;
            //--------------------------------

            double Ukonmin;
            double Ukonmax;
            double Vkonmin;
            double Vkonmax;
            double Upercentkon;
            double Mkon1;
            double Ukon1;

            //Для подтоварной воды конечной
            double UkonminH2O;
            double UkonmaxH2O;
            double VkonminH2O;
            double VkonmaxH2O;
            double V2H2O;
            double V2RH2O;
            double UpercentkonH2O;
            //--------------------------------

            List<calibration> calibrationList = new List<calibration>();
            calibrationList = db.calibration.Where(f => f.filialid == filial2).OrderBy(h => h.tankid).ThenBy(h => h.oillevel).ToList();

            if (filial2 == 1)
            {
                LastTanksResultMoz MOZ = new LastTanksResultMoz();
                MOZ = db2.LastTanksResultMoz.FirstOrDefault(f => f.tankid == rezer2);
                if (MOZ == null)
                {
                    H.dat = null;
                    H.densreal = null;
                    H.dens20 = null;
                    H.tempreal = null;
                    H.water = null;
                    H.saltmg = null;
                    H.mechan = null;
                }
                else
                {
                    H.dat = MOZ.endsampledt;
                    H.densreal = MOZ.densreal;
                    H.dens20 = MOZ.dens20;
                    H.tempreal = Convert.ToDouble(MOZ.tempreal);
                    H.water = MOZ.water;
                    H.saltmg = MOZ.saltmg;
                    H.mechan = MOZ.mechan;
                }
            }

            if (filial2 == 2)
            {
                LastTanksResultPol MOZ = new LastTanksResultPol();
                MOZ = db2.LastTanksResultPol.FirstOrDefault(j => j.tankid == rezer2);
                if (MOZ == null)
                {
                    H.dat = null;
                    H.densreal = null;
                    H.dens20 = null;
                    H.tempreal = null;
                    H.water = null;
                    H.saltmg = null;
                    H.mechan = null;
                }
                else
                {
                    H.dat = MOZ.sampledt;
                    H.densreal = MOZ.densreal;
                    H.dens20 = MOZ.dens20;
                    H.tempreal = Convert.ToDouble(MOZ.tempreal);
                    H.water = MOZ.water;
                    H.saltmg = MOZ.saltmg;
                    H.mechan = MOZ.mechan;
                }
            }

            if (calibrationList.Where(h => h.tankid == rezer2).Count() == 0)
            {
                ViewBag.V1 = "данных нет!";
                ViewBag.V2 = "даных нет!";
                ViewBag.Mnach = "данных нет!";
                ViewBag.Mkon = "данных нет!";
                ViewBag.RV = "данных нет!";
                ViewBag.RM = "данных нет!";
            }
            else if (calibrationList.LastOrDefault(g => g.tankid == rezer2).oillevel < Convert.ToDouble(Unach1))
            {
                ViewBag.V1 = "Большое значение!";
                ViewBag.V2 = "Большое значение!";
                ViewBag.Mnach1 = "Большое значение!";
                ViewBag.Mkon1 = "Большое значение!";
                ViewBag.RV1 = "Большое значение!";
                ViewBag.Ukon1 = "Большое значение!";
                ViewBag.MNet = "Большое значение!";
            }
            else
            {
                //-------Рассчет объема нефти начального---------------------------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(Unach1)) == null)
                {
                    Unachmin = 0;
                }
                else
                {
                    Unachmin = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(Unach1)).oillevel;
                }
                Unachmax = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oillevel > Convert.ToDouble(Unach1)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(Unach1)) == null)
                {
                    Vnachmin = 0;
                }
                else
                {
                    Vnachmin = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(Unach1)).oilvolume;
                }
                Vnachmax = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oillevel > Convert.ToDouble(Unach1)).oilvolume;
                Upercentnach = (Convert.ToDouble(Unach1) - Unachmin) / (Unachmax - Unachmin);
                V1 = Vnachmin + (Vnachmax - Vnachmin) * Upercentnach; //по градуировочной таблице
                //V1R = V1 * (1 + (2 * Convert.ToDouble(Kst1) + Convert.ToDouble(Ksr1)) * (Convert.ToDouble(Tnach1) - 20)); //рассчитано по формуле

                //---------Рассчет объема подтоварной воды начальной-------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UnachH2O1)) == null)
                {
                    UnachminH2O = 0;
                }
                else
                {
                    UnachminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UnachH2O1)).oillevel;
                }
                UnachmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oillevel > Convert.ToDouble(UnachH2O1)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UnachH2O1)) == null | Convert.ToDouble(UnachH2O1) == 0)
                {
                    VnachminH2O = 0;
                }
                else
                {
                    VnachminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UnachH2O1)).oilvolume;
                }
                VnachmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oillevel > Convert.ToDouble(UnachH2O1)).oilvolume;
                UpercentnachH2O = (Convert.ToDouble(UnachH2O1) - UnachminH2O) / (UnachmaxH2O - UnachminH2O);
                V1H2O = VnachminH2O + (VnachmaxH2O - VnachminH2O) * UpercentnachH2O;
                //V1RH2O = V1H2O * (1 + (2 * Convert.ToDouble(Kst1) + Convert.ToDouble(Ksr1)) * (Convert.ToDouble(Tnach1) - 20));

                V1RN = (Math.Round(V1, 3) - Math.Round(V1H2O, 3)) * (1 + (2 * Convert.ToDouble(Kst1) + Convert.ToDouble(Ksr1)) * (Convert.ToDouble(Tnach1) - 20)); // Рассчет начального объема чистой нефти без подтоварной воды

                //------------Рассчет объема подтоварной воды конечной-----------------------------------------------------------------------
                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UkonH2O1)) == null)
                {
                    UkonminH2O = 0;
                }
                else
                {
                    UkonminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UkonH2O1)).oillevel;
                }
                UkonmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oillevel > Convert.ToDouble(UkonH2O1)).oillevel;
                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UkonH2O1)) == null | Convert.ToDouble(UkonH2O1) == 0)
                {
                    VkonminH2O = 0;
                }
                else
                {
                    VkonminH2O = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oillevel <= Convert.ToDouble(UkonH2O1)).oilvolume;
                }
                VkonmaxH2O = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oillevel > Convert.ToDouble(UkonH2O1)).oilvolume;
                UpercentkonH2O = (Convert.ToDouble(UkonH2O1) - UkonminH2O) / (UkonmaxH2O - UkonminH2O);
                V2H2O = VkonminH2O + (VkonmaxH2O - VkonminH2O) * UpercentkonH2O;
                V2RH2O = V2H2O * (1 + (Convert.ToDouble(Kst1) + Convert.ToDouble(Ksr1)) * (Convert.ToDouble(Tkon1) - 20));

                //V2RN = (Math.Round(V2, 3) - Math.Round(V2H2O, 3)) * (1 + (2 * Convert.ToDouble(Kst) + Convert.ToDouble(Ksr)) * (Convert.ToDouble(Tkon) - 20)); //Расчет конечного объема читой нефти без подтоварной воды

                //--------------------------------------------------------------------------------------------------------------------
                //V1RN = (Math.Round(V1, 3) - Math.Round(V1H2O, 3)) * (1 + (2 * Convert.ToDouble(Kst1) + Convert.ToDouble(Ksr1)) * (Convert.ToDouble(Tnach1) - 20)); // Рассчет конечного объема нефти без подтоварной воды

                Mnach1 = Convert.ToDouble(Pnach1) * Math.Round(V1RN, 3) / 1000; // Масса нефти начальная

                RMB = 100 * Convert.ToDouble(RM1) / (100 - Convert.ToDouble(Bal1)); //Рассчет массы(брутто) нефти

                Mkon1 = Convert.ToDouble(RMB) + Math.Round(Mnach1, 3); //Рассчет массы конечной

                V2RN = 1000 * Mkon1 / Convert.ToDouble(Pkon1); //Рассчет объема нефти чистой конечной

                V2 = V2RN / (1 + (2 * Convert.ToDouble(Kst1) + Convert.ToDouble(Ksr1)) * (Convert.ToDouble(Tkon1) - 20)) + V2H2O;

                if (calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oilvolume <= Convert.ToDouble(V2)) == null | calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oilvolume > Convert.ToDouble(V2)) == null)
                {
                    ViewBag.V1 = "Большое значение!";
                    ViewBag.V2 = "Большое значение!";
                    ViewBag.Mnach1 = "Большое значение!";
                    ViewBag.Mkon1 = "Большое значение!";
                    ViewBag.RV1 = "Большое значение!";
                    ViewBag.Ukon1 = "Большое значение!";
                    ViewBag.MNet = "Большое значение!";
                }
                else
                {
                    Vkonmin = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oilvolume <= Convert.ToDouble(V2)).oilvolume;
                    Vkonmax = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oilvolume > Convert.ToDouble(V2)).oilvolume;
                    Upercentkon = (V2 - Vkonmin) / (Vkonmax - Vkonmin);

                    Ukonmin = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oilvolume <= Convert.ToDouble(V2)).oillevel;
                    Ukonmax = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oilvolume > Convert.ToDouble(V2)).oillevel;
                    Ukon1 = Ukonmin + (Ukonmax - Ukonmin) * Upercentkon;

                    //Mkon1 = Convert.ToDouble(RM1) + Mnach1;
                    //V2 = Mkon1 / Convert.ToDouble(Pkon1) * 1000;
                    //V2R = V2 * (1 + 0.0000375 * (Convert.ToDouble(Tkon1) - 20));
                    RV = V2RN - V1RN;

                    //Vkonmin = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oilvolume <= Convert.ToDouble(V2R)).oilvolume;
                    //Vkonmax = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oilvolume > Convert.ToDouble(V2R)).oilvolume;
                    //Ukonmin = calibrationList.LastOrDefault(g => g.tankid == rezer2 & g.oilvolume <= Convert.ToDouble(V2R)).oillevel;
                    //Ukonmax = calibrationList.FirstOrDefault(g => g.tankid == rezer2 & g.oilvolume > Convert.ToDouble(V2R)).oillevel;
                    //Upercentkon = (V2R - Vkonmin) / (Vkonmax - Vkonmin);
                    //Ukon1 = Ukonmin + (Ukonmax - Ukonmin) * Upercentkon; 


                    ViewBag.V1 = Math.Round(V1RN, 3);
                    ViewBag.V2 = Math.Round(V2RN, 3);
                    ViewBag.Mnach1 = Math.Round(Mnach1, 3);
                    ViewBag.Mkon1 = Math.Round(Mkon1, 3);
                    ViewBag.RV1 = Math.Round(RV, 0);
                    //ViewBag.RM = RM;
                    ViewBag.Ukon1 = Math.Round(Ukon1, 0);
                    ViewBag.MNet = Math.Round((Convert.ToDouble(RM1) + (Convert.ToDouble(RM1) * Convert.ToDouble(Bal1) / 100)), 0);
                }
            }
            return PartialView();
        }
        //-----Формирование отчета в EXCEL----------------------------------------------------------------------

        public ActionResult Export(int filial1, string datinv)
        {
            DateTime datinvAct = Convert.ToDateTime(datinv);

            List<TankInv> Vod = new List<TankInv>();
            Vod = db.TankInv.Where(p => p.Filial == filial1).Where(f => f.Data == datinvAct).OrderBy(g => g.type).ThenBy(g => g.Rezer).ToList();
            var VodGr = Vod.GroupBy(h => h.type); //список сгруппированный по полю type

            List<trl_tank> trl = new List<trl_tank>();
            trl = db.trl_tank.Where(p => p.FilialID == filial1).OrderBy(p => p.TypeID).ThenBy(p => p.TankID).ToList();

            List<int> typ = new List<int>();
            foreach (var t in trl)
            {
                {
                    typ.Add(t.TypeID);
                }
            }

            List<int> typDyst = new List<int>();

            //---Список типов резервуаров---------------------
            typDyst = typ.Distinct().ToList();
            //-------------------------------------------------

            TankInfoType TanType = new TankInfoType();
            foreach (var p in Vod)
            {
                TanType.tankid = p.Rezer;
                TanType.date = p.Data;
                TanType.filial = p.Filial;
                TanType.level = Convert.ToDouble(p.Urov);
                TanType.temp = Convert.ToDouble(p.Temp);
                if (trl.FirstOrDefault(g => g.TankID == p.Rezer) == null)
                {
                    TanType.type = 777;
                }
                else
                {
                    TanType.type = trl.FirstOrDefault(g => g.TankID == p.Rezer).TypeID;
                }

            }
            string MOL = "";
            List<Podpisanty> Podpis = new List<Podpisanty>();
            Podpis = db.Podpisanty.ToList();

            Podpisanty Predsed = db.Podpisanty.FirstOrDefault(g => g.Nazn == "Председатель комиссии");
            if (db.Podpisanty.FirstOrDefault(g => g.Nazn == "МОЛ") == null)
            {
                MOL = "";
            }
            else
            {
                MOL = db.Podpisanty.FirstOrDefault(g => g.Nazn == "МОЛ").FIO;
            }

            //---Задаем параметры листа------------------------------------------------------------

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Лист1");
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            worksheet.PageSetup.FitToPages(1, 1);

            //Шапка отчета//
            //worksheet.Columns().AdjustToContents();
            string chleny = "";


            if (filial1 == 1)
            {
                worksheet.Cell("G" + 1).Value = "Акт инвентаризации нефти в резервуарах филиала ЛПДС 'Мозырь'";
            }
            else
            {
                worksheet.Cell("G" + 1).Value = "Акт инвентаризации нефти в резервуарах филиала ЛПДС 'Новополоцк'";
            }

            worksheet.Cell("A" + 2).Value = "Председатель комиссии - " + Predsed.Doljnost.Trim() + " " + Predsed.FIO.Trim();
            foreach (var i in Podpis)
            {
                chleny = chleny + i.Doljnost.Trim() + " " + i.FIO.Trim() + ", ";
            }
            worksheet.Cell("A" + 3).Value = "Члены комиссии - " + chleny;
            worksheet.Cell("A" + 4).Value = "составили настоящий акт в том, что по состоянию на 24 часа московского времени " + Convert.ToDateTime(datinv).ToShortDateString() + "было установлено наличие нефти следующего количества";
            worksheet.Cell("E" + 1).Style.Font.FontSize = 12;

            worksheet.Cell("C" + 2).Style.Font.FontSize = 14;
            worksheet.Cell("C" + 1).Style.Font.FontSize = 14;

            worksheet.Cell("G" + 1).Style.Font.FontSize = 14;
            worksheet.Cell("G" + 1).Style.Font.Bold = true;

            //----Этот стиль шапки описан в общем блоке стиле для шапки таблицы-----------------
            //worksheet.Cell("A" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("B" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("C" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("D" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("E" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("F" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("G" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("H" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("I" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("J" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("K" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("L" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("M" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("N" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("O" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("P" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("Q" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("R" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("S" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("T" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("U" + 7).Style.Alignment.WrapText = true;
            //worksheet.Cell("V" + 7).Style.Alignment.WrapText = true;

            //worksheet.Cell("A" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("B" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("C" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("D" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("E" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("F" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("G" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("H" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("I" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("J" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("K" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("L" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("M" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("N" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("O" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("P" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("Q" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("R" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("S" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("T" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("U" + 6).Style.Alignment.WrapText = true;
            //worksheet.Cell("V" + 6).Style.Alignment.WrapText = true;

            //создадим заголовки у столбцов
            worksheet.Cell("A" + 6).Value = "Номер резервуара";
            worksheet.Cell("B" + 6).Value = "Общий уровень нефти";
            worksheet.Cell("B" + 9).Value = "мм";
            worksheet.Cell("C" + 6).Value = "Уровень подтоварной воды";
            worksheet.Cell("C" + 9).Value = "мм";
            worksheet.Cell("D" + 6).Value = "Общий объем";
            worksheet.Cell("D" + 9).Value = "м3";
            worksheet.Cell("E" + 6).Value = "Объем подтоварной воды и донных отложений";
            worksheet.Cell("E" + 9).Value = "м3";
            worksheet.Cell("F" + 6).Value = "Объем нефти";
            worksheet.Cell("F" + 9).Value = "м3";
            worksheet.Cell("G" + 6).Value = "Средняя темп.нефти";
            worksheet.Cell("G" + 9).Value = "С";
            worksheet.Cell("H" + 6).Value = "Плотность нефти при ср.темп.";
            worksheet.Cell("H" + 9).Value = "кг/м3";
            worksheet.Cell("I" + 6).Value = "Масса брутто нефти";
            worksheet.Cell("I" + 9).Value = "т";
            worksheet.Cell("J" + 7).Value = "Вода";
            worksheet.Cell("J" + 9).Value = "%";
            worksheet.Cell("K" + 7).Value = "Соль";
            worksheet.Cell("K" + 9).Value = "%";
            worksheet.Cell("L" + 7).Value = "Мех.прим.";
            worksheet.Cell("L" + 9).Value = "%";
            worksheet.Cell("M" + 9).Value = "%";
            worksheet.Cell("N" + 9).Value = "т";
            worksheet.Cell("O" + 6).Value = "Масса нефти нетто";
            worksheet.Cell("O" + 9).Value = "т";
            worksheet.Cell("P" + 8).Value = "Нижний нормативный уровень";
            worksheet.Cell("P" + 9).Value = "мм"; ;
            worksheet.Cell("Q" + 8).Value = "Объем по нижнему нормативному уровню";
            worksheet.Cell("Q" + 9).Value = "м3";
            worksheet.Cell("R" + 8).Value = "Масса брутто";
            worksheet.Cell("R" + 9).Value = "т";
            worksheet.Cell("S" + 8).Value = "Масса нетто";
            worksheet.Cell("S" + 9).Value = "т";
            worksheet.Cell("T" + 7).Value = "Товарный";
            worksheet.Cell("T" + 8).Value = "Масса нетто";
            worksheet.Cell("T" + 9).Value = "т";

            worksheet.Cell("J" + 6).Value = "Содержание баласта";
            worksheet.Cell("P" + 6).Value = "В том числе остатки в резервуарах";
            worksheet.Cell("M" + 7).Value = "Всего";
            worksheet.Cell("P" + 7).Value = "Технологический";

            worksheet.Cell("E6").Style.Font.FontSize = 8;
            worksheet.Cell("P8").Style.Font.FontSize = 8;
            worksheet.Cell("Q8").Style.Font.FontSize = 8;

            worksheet.Range("A6:A9").Column(1).Merge();
            worksheet.Range("B6:B8").Column(1).Merge();
            worksheet.Range("C6:C8").Column(1).Merge();
            worksheet.Range("D6:D8").Column(1).Merge();
            worksheet.Range("E6:E8").Column(1).Merge();
            worksheet.Range("F6:F8").Column(1).Merge();
            worksheet.Range("G6:G8").Column(1).Merge();
            worksheet.Range("H6:H8").Column(1).Merge();
            worksheet.Range("N6:N8").Column(1).Merge();
            worksheet.Range("I6:I8").Column(1).Merge();
            worksheet.Range("J7:J8").Column(1).Merge();
            worksheet.Range("K7:K8").Column(1).Merge();
            worksheet.Range("L7:L8").Column(1).Merge();
            worksheet.Range("O6:O8").Column(1).Merge();
            worksheet.Range("J6:N6").Row(1).Merge();
            worksheet.Range("M7:N7").Row(1).Merge();
            worksheet.Range("P6:T6").Row(1).Merge();
            worksheet.Range("P7:S7").Row(1).Merge();

            double VNeftItog = 0;
            double PSred = 0;
            double MassaBalItog = 0;
            double BalProcSred = 0;
            double BalTonnItog = 0;
            double MassaNettoItog = 0;
            double MassaBalMinItog = 0;
            double MassaNettoMinItog = 0;


            for (int ii = 0; ii < Vod.Count; ii++)

            {
                VNeftItog = VNeftItog + Convert.ToDouble(Vod[ii].VNeft);
                PSred = Convert.ToDouble(Vod[ii].P) + PSred;
                MassaBalItog = MassaBalItog + Convert.ToDouble(Vod[ii].MassaBrutto);
                BalProcSred = BalProcSred + Convert.ToDouble(Vod[ii].BalProc);
                BalTonnItog = BalTonnItog + Convert.ToDouble(Vod[ii].BalTonn);
                MassaNettoItog = MassaNettoItog + Convert.ToDouble(Vod[ii].MassaNetto);
                MassaBalMinItog = MassaBalMinItog + Convert.ToDouble(Vod[ii].MBalMin);
                MassaNettoMinItog = MassaNettoMinItog + Convert.ToDouble(Vod[ii].MNettoMin);

            }

            int uu = -1;

            var groupss = from p in Vod
                          group p by p.type;
            foreach (var gg in groupss)
            {
                int kol = 0;
                double VNeftItog1 = 0;
                double PSred1 = 0;
                double MassaBalItog1 = 0;
                double BalProcSred1 = 0;
                double BalTonnItog1 = 0;
                double MassaNettoItog1 = 0;
                double MassaBalMinItog1 = 0;
                double MassaNettoMinItog1 = 0;

                foreach (var pp in gg)
                {
                    uu++;
                    worksheet.Cell("A" + (uu + 10)).Value = pp.Rezer;
                    worksheet.Cell("B" + (uu + 10)).Value = pp.Urov;
                    worksheet.Cell("C" + (uu + 10)).Value = pp.UrovH2O;
                    worksheet.Cell("D" + (uu + 10)).Value = pp.V;
                    worksheet.Cell("E" + (uu + 10)).Value = pp.VH2O;
                    worksheet.Cell("F" + (uu + 10)).Value = pp.VNeft;
                    worksheet.Cell("G" + (uu + 10)).Value = pp.Temp;
                    worksheet.Cell("H" + (uu + 10)).Value = pp.P;
                    worksheet.Cell("I" + (uu + 10)).Value = pp.MassaBrutto;
                    worksheet.Cell("J" + (uu + 10)).Value = pp.H2O;
                    worksheet.Cell("K" + (uu + 10)).Value = pp.Salt;
                    worksheet.Cell("L" + (uu + 10)).Value = pp.Meh;
                    worksheet.Cell("M" + (uu + 10)).Value = pp.BalProc;
                    worksheet.Cell("N" + (uu + 10)).Value = pp.BalTonn;
                    worksheet.Cell("O" + (uu + 10)).Value = pp.MassaNetto;
                    worksheet.Cell("P" + (uu + 10)).Value = pp.HMim;
                    worksheet.Cell("Q" + (uu + 10)).Value = pp.VMin;
                    worksheet.Cell("R" + (uu + 10)).Value = pp.MBalMin;
                    worksheet.Cell("S" + (uu + 10)).Value = pp.MNettoMin;
                    worksheet.Cell("T" + (uu + 10)).Value = null;

                    kol++;
                    VNeftItog1 = VNeftItog1 + Convert.ToDouble(pp.VNeft);
                    PSred1 = Convert.ToDouble(pp.P) + PSred;
                    MassaBalItog1 = MassaBalItog1 + Convert.ToDouble(pp.MassaBrutto);
                    BalProcSred1 = BalProcSred1 + Convert.ToDouble(pp.BalProc);
                    BalTonnItog1 = BalTonnItog1 + Convert.ToDouble(pp.BalTonn);
                    MassaNettoItog1 = MassaNettoItog1 + Convert.ToDouble(pp.MassaNetto);
                    MassaBalMinItog1 = MassaBalMinItog1 + Convert.ToDouble(pp.MBalMin);
                    MassaNettoMinItog1 = MassaNettoMinItog1 + Convert.ToDouble(pp.MNettoMin);

                }
                uu++;
                worksheet.Cell("A" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("B" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("C" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("D" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("E" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("F" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("G" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("H" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("I" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("J" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("K" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("L" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("M" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("N" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("O" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("P" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("Q" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("R" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("S" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;
                worksheet.Cell("T" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.MossGreen;

                worksheet.Cell("A" + (uu + 10)).Value = "ИТОГО ПО ГРУППЕ: ";
                worksheet.Cell("F" + (uu + 10)).Value = VNeftItog1;
                worksheet.Cell("H" + (uu + 10)).Value = PSred1 / kol;
                worksheet.Cell("I" + (uu + 10)).Value = MassaBalItog1;
                worksheet.Cell("N" + (uu + 10)).Value = BalTonnItog1;
                worksheet.Cell("O" + (uu + 10)).Value = MassaNettoItog1;
                worksheet.Cell("S" + (uu + 10)).Value = MassaNettoMinItog1;

            }
            uu++;
            worksheet.Cell("A" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("B" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("C" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("D" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("E" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("F" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("G" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("H" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("I" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("J" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("K" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("L" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("M" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("N" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("O" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("P" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("Q" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("R" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("S" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;
            worksheet.Cell("T" + (uu + 10)).Style.Fill.BackgroundColor = XLColor.BabyBlue;

            worksheet.Cell("A" + (uu + 10)).Value = "ИТОГО:";
            worksheet.Cell("F" + (uu + 10)).Value = VNeftItog;
            worksheet.Cell("H" + (uu + 10)).Value = PSred / Vod.Count;
            worksheet.Cell("I" + (uu + 10)).Value = MassaBalItog;
            worksheet.Cell("N" + (uu + 10)).Value = BalTonnItog;
            worksheet.Cell("O" + (uu + 10)).Value = MassaNettoItog;
            worksheet.Cell("S" + (uu + 10)).Value = MassaNettoMinItog;
            worksheet.Cell("T" + (uu + 10)).Value = MassaNettoItog - MassaNettoMinItog;
            //---------------------------------------------------------

            //--------Этот стиль описан в одном блоке с остальными стилями шапки таблицы-------------------------
            //worksheet.Cell("A" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("B" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("C" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("D" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("E" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("F" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("G" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("H" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("I" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("J" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("K" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("L" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("M" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("N" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("O" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("P" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("Q" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("R" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("S" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("T" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("U" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("V" + 7).Style.Fill.BackgroundColor = XLColor.MossGreen;

            //worksheet.Cell("A" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("B" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("C" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("D" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("E" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("F" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("G" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("H" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("I" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("J" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("K" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("L" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("M" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("N" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("O" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("P" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("Q" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("R" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("S" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("T" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("U" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;
            //worksheet.Cell("V" + 6).Style.Fill.BackgroundColor = XLColor.MossGreen;



            //----------Заголовки в шапке по центру (закоментировал и описал этот стиль в одном блоке с остальными стилями)----
            //worksheet.Cell("A7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("B7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("C7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("D7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("E7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("F7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("G7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("H7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("I7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("N7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("O7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("S7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("V7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            //worksheet.Cell("I6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("J6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("K6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("L6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("M6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("O6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("P6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("Q6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("R6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("S6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("T6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("U6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //worksheet.Cell("V6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var rngTable = worksheet.Range("A6:T" + (uu + 10));
            rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;


            //---Здесь описаны стили для шапки таблиц-----------------------------------//
            var rngTable111 = worksheet.Range("A6:T" + 9);
            rngTable111.Style.Border.RightBorder = XLBorderStyleValues.Medium;
            rngTable111.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            rngTable111.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
            rngTable111.Style.Border.TopBorder = XLBorderStyleValues.Medium;
            rngTable111.Style.Fill.BackgroundColor = XLColor.MossGreen;
            rngTable111.Style.Alignment.WrapText = true;
            rngTable111.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            //------Вывод информации под таблицей где чья нефть ---------------------------------------------

            worksheet.Cell("C" + (uu + 12)).Value = "Собственность ОАО 'ОАО Гомельтранснефть Дружба', находящаяся в резервуарах филиала 'ЛПДС Мозырь'";
            worksheet.Cell("C" + (uu + 13)).Value = "Собственность РУП 'Производственное объединение 'Белоруснефть'";
            worksheet.Cell("C" + (uu + 14)).Value = "Ресурсы ОАО 'Мозырский НПЗ', хранимые в арендованных резервуарах";
            worksheet.Cell("C" + (uu + 15)).Value = "Ресурсы ОАО 'Мозырский НПЗ', хранимые по договору хранения";
            worksheet.Cell("C" + (uu + 16)).Value = "Ответственное хранение:";

            //-------Вывод под таблицей сколько чьей-то нефти------------------------------------------------------------
            worksheet.Cell("O" + (uu + 12)).Value = "0,0 тонн нетто";
            worksheet.Cell("O" + (uu + 13)).Value = "0,0 тонн нетто";
            worksheet.Cell("O" + (uu + 14)).Value = "0,0 тонн нетто";
            worksheet.Cell("O" + (uu + 15)).Value = "0,0 тонн нетто";
            worksheet.Cell("O" + (uu + 16)).Value = "0,0 тонн нетто";

            //------Вывод подписантов внизу------------------------------------------------------
            worksheet.Cell("A" + (uu + 20)).Value = "Председатель комиссии________________" + Predsed.FIO;
            worksheet.Cell("H" + (uu + 20)).Value = "Члены комиссии________________";
            worksheet.Cell("A" + (uu + 22)).Value = "Материально ответственное лицо________________" + MOL;
            //---------------------------------------------------------------------------------

            //var col1 = worksheet.Column("A");
            //col1.AdjustToContents();

            //var col2 = worksheet.Column("B");
            //col2.Width = 40;

            //var col3 = worksheet.Column("C");
            //col3.Width = 40;

            //var col4 = worksheet.Column("D");
            //col4.Width = 14;

            //var col5 = worksheet.Column("E");
            //col5.Width = 18;

            //var col6 = worksheet.Column("F");
            //col5.Width = 18;

            //worksheet.Columns().AdjustToContents();

            //worksheet.Column(1).Style.Alignment.WrapText = true;
            //worksheet.Column(2).Style.Alignment.WrapText = true;
            //worksheet.Column(3).Style.Alignment.WrapText = true;
            //worksheet.Column(4).Style.Alignment.WrapText = true;
            //worksheet.Column(6).Style.Alignment.WrapText = true;

            // вернем пользователю файл без сохранения его на сервере

            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
            }
        }

        //------------------------------------------------------------------------------------------------------

        // Редактирование члена комиссии//

        public ActionResult InventEdit(int ID)
        {
            TankInv TanI = new TankInv();
            TanI = db.TankInv.FirstOrDefault(a => a.Id == ID);
            return PartialView(TanI);
        }
        //-------------------------------//

        //Сохранение редактирования члена комиссии------------//
        [HttpPost]
        public ActionResult InventEditSave(int ID, string InvRezer, string InvUrov, string InvUrovH2O, string InvV, string InvVH2O, string InvVNeft, string InvTemp, string InvP, string InvMassaBrutto, string InvH2O, string InvSalt, string InvMeh, string InvBalProc, string InvBalTonn, string InvMassaNetto, string InvHMim, string InvVMin, string InvMNettoMin, string InvVTeh)
        {
            try
            {
                TankInv TanKINVE = new TankInv();
                TanKINVE = db.TankInv.FirstOrDefault(s => s.Id == ID);

                TanKINVE.Rezer = Convert.ToInt32(InvRezer);
                TanKINVE.Urov = Convert.ToDecimal(InvUrov);
                TanKINVE.UrovH2O = Convert.ToDecimal(InvUrovH2O);
                TanKINVE.V = Convert.ToDecimal(InvV);
                TanKINVE.VH2O = Convert.ToDecimal(InvVH2O);
                TanKINVE.VNeft = Convert.ToDecimal(InvVNeft);
                TanKINVE.Temp = Convert.ToDecimal(InvTemp);
                TanKINVE.P = Convert.ToDecimal(InvP);
                TanKINVE.MassaBrutto = Convert.ToDecimal(InvMassaBrutto);
                TanKINVE.H2O = Convert.ToDecimal(InvH2O);
                TanKINVE.Salt = Convert.ToDecimal(InvSalt);
                TanKINVE.Meh = Convert.ToDecimal(InvMeh);
                TanKINVE.BalProc = Convert.ToDecimal(InvBalProc);
                TanKINVE.BalTonn = Convert.ToDecimal(InvBalTonn);
                TanKINVE.MassaNetto = Convert.ToDecimal(InvMassaNetto);
                TanKINVE.HMim = Convert.ToDecimal(InvHMim);
                TanKINVE.VMin = Convert.ToDecimal(InvVMin);
                TanKINVE.MNettoMin = Convert.ToDecimal(InvMNettoMin);
                TanKINVE.VTeh = Convert.ToDecimal(InvVTeh);

                db.SaveChanges();

                ViewBag.Message = "Инвентаризация успешно изменена";

            }
            catch (Exception e)
            {
                ViewBag.Message = "Ошибка. Текст ошибки: " + e.ToString();
            }

            return PartialView();
        }
        //-----------------------------//
        public string GetUrovH2O(int ID, int IDFilial, int InvRezerEdit, string InvUrovEdit, string InvUrovH2OEdit, string InvVEdit, string InvVH2OEdit, string InvVNeftEdit, string InvTempEdit, string InvPEdit, 
        string InvMassaBruttoEdit, string InvH2OEdit, string InvSaltEdit, string InvMehEdit, string InvBalProcEdit, string InvBalTonnEdit, string InvMassaNettoEdit, string InvHMimEdit, 
        string InvVMinEdit, string InvMNettoMinEdit, string InvMBalMinEdit)
        {
            //---------------------------------------------------------------------------- 
            double VEdit = Convert.ToDouble(InvVEdit);
            double InvUrov = Convert.ToDouble(InvUrovEdit);
            double InvUrovH2O = Convert.ToDouble(InvUrovH2OEdit);
            double InvUrovNeft = InvUrov - InvUrovH2O;
            double VH2OEdit = Convert.ToDouble(InvVH2OEdit);
            double VNeftEdit = Convert.ToDouble(InvVNeftEdit);
            double MassaBruttoEdit = Convert.ToDouble(InvMassaBruttoEdit);
            double BalProcEdit = Convert.ToDouble(InvBalProcEdit);
            double BalTonnEdit = Convert.ToDouble(InvBalTonnEdit);
            double MassaNettoEdit = Convert.ToDouble(InvMassaNettoEdit);
            double UminEdit;
            double UmaxEdit;
            double UpercentEdit;
            double VminEdit;
            double VmaxEdit;            
                        
            double UminEditH2O;
            double UmaxEditH2O;
            double UpercentEditH2O;
            double VminEditH2O;
            double VmaxEditH2O;
            double VNeft;
            double MassaBrutto;
            double BalProc;
            double BalTonn;
            double MassaNetto;
            double HMin;
            double VMin;

            List<calibration> calibrationListEdit = new List<calibration>();
            calibrationListEdit = db.calibration.Where(f => f.filialid == IDFilial & f.tankid == InvRezerEdit).ToList();

            //Рассчет объема общего
            if (calibrationListEdit == null)
            {
                VEdit = 0;                
            }
            else
            {
                if (calibrationListEdit.LastOrDefault(g=>g.oillevel <= Convert.ToDecimal(InvUrovEdit)) == null)
                {
                    UminEdit = 0;
                }
                else
                {
                    UminEdit = calibrationListEdit.LastOrDefault(g=>g.oillevel <= Convert.ToDouble(InvUrovEdit)).oillevel;
                }
                UmaxEdit = calibrationListEdit.FirstOrDefault(j=>j.oillevel > Convert.ToDouble(InvUrovEdit)).oillevel;

                if (calibrationListEdit.LastOrDefault(g => g.oillevel <= Convert.ToDouble(InvUrovEdit)) == null)
                {
                    VminEdit = 0;
                }
                else
                {
                    VminEdit = calibrationListEdit.LastOrDefault(g => g.oillevel <= Convert.ToDouble(InvUrovEdit)).oilvolume;
                }
                VmaxEdit = calibrationListEdit.FirstOrDefault(j => j.oillevel > Convert.ToDouble(InvUrovEdit)).oilvolume;
                UpercentEdit = (Convert.ToDouble(InvUrovEdit) - UminEdit) / (UmaxEdit - UminEdit);
                VEdit = VminEdit + (VmaxEdit - VminEdit) * UpercentEdit;
                }

            //--------Рассчет объема подтоварной воды----------------------------------------------------------------
            if (calibrationListEdit == null)
            {
                VH2OEdit = 0;
            }
            else
            {
                if (calibrationListEdit.LastOrDefault(g => g.oillevel <= Convert.ToDouble(InvUrovH2OEdit)) == null)
                {
                    UminEditH2O = 0;
                }
                else
                {
                    UminEditH2O = calibrationListEdit.LastOrDefault(g => g.oillevel <= Convert.ToDouble(InvUrovH2OEdit)).oillevel;
                }
                UmaxEditH2O = calibrationListEdit.FirstOrDefault(j => j.oillevel > Convert.ToDouble(InvUrovH2OEdit)).oillevel;

                if (calibrationListEdit.LastOrDefault(g => g.oillevel <= Convert.ToDouble(InvUrovH2OEdit)) == null)
                {
                    VminEditH2O = 0;
                }
                else
                {
                    VminEditH2O = calibrationListEdit.LastOrDefault(g => g.oillevel <= Convert.ToDouble(InvUrovH2OEdit)).oilvolume;
                }
                VmaxEditH2O = calibrationListEdit.FirstOrDefault(j => j.oillevel > Convert.ToDouble(InvUrovH2OEdit)).oilvolume;
                UpercentEditH2O = (Convert.ToDouble(InvUrovH2OEdit) - UminEditH2O) / (UmaxEditH2O - UminEditH2O);
                VH2OEdit = VminEditH2O + (VmaxEditH2O - VminEditH2O) * UpercentEditH2O;
            }
            //-------------------------------------------------------------------------------
            VNeft = VEdit - VH2OEdit; // Объем нефти;
            MassaBrutto = VNeft * Convert.ToDouble(InvPEdit) / 1000; // Масса нефти брутто;
            BalProc = Convert.ToDouble(InvH2OEdit) + Convert.ToDouble(InvSaltEdit) + Convert.ToDouble(InvMehEdit); //Рассчет баласта в процентах;
            BalTonn = MassaBrutto * BalProc / 100; //Рассчет балласта в тоннах;
            MassaNetto = MassaBrutto - BalTonn;
            HMin = Convert.ToDouble(InvHMimEdit);
            VMin = Convert.ToDouble(InvVMinEdit);

            TankInv TanEdit = new TankInv();
            TanEdit.Id = ID;
            TanEdit.Filial = IDFilial;
            TanEdit.Rezer = InvRezerEdit;
            TanEdit.Urov = Convert.ToDecimal(InvUrov);
            TanEdit.UrovH2O = Convert.ToDecimal(InvUrovH2O);
            TanEdit.V = Math.Round(Convert.ToDecimal(VEdit),3);
            TanEdit.VH2O = Math.Round(Convert.ToDecimal(VH2OEdit),3);
            TanEdit.UrovNeft = Convert.ToDecimal(InvUrovNeft);
            TanEdit.VNeft = Math.Round(Convert.ToDecimal(VNeft), 3);
            TanEdit.Temp = Convert.ToDecimal(InvTempEdit);
            TanEdit.P = Convert.ToDecimal(InvPEdit);
            TanEdit.MassaBrutto = Math.Round(Convert.ToDecimal(MassaBrutto), 3);
            TanEdit.H2O = Convert.ToDecimal(InvH2OEdit);
            TanEdit.Salt = Convert.ToDecimal(InvSaltEdit);
            TanEdit.Meh = Convert.ToDecimal(InvMehEdit);
            TanEdit.BalProc = Math.Round(Convert.ToDecimal(BalProc), 3);
            TanEdit.BalTonn = Math.Round(Convert.ToDecimal(BalTonn), 3);
            TanEdit.MassaNetto = Math.Round(Convert.ToDecimal(MassaNetto), 3);
            TanEdit.HMim = Convert.ToDecimal(HMin);
            TanEdit.VMin = Convert.ToDecimal(VMin);
            TanEdit.MBalMin = Convert.ToDecimal(VMin) * Convert.ToDecimal(InvPEdit) / 1000;
            TanEdit.MNettoMin = TanEdit.MBalMin - (TanEdit.MBalMin * Math.Round(Convert.ToDecimal(BalProc), 3) / 100);

            ViewBag.TanEdit = TanEdit;
            //ViewBag.TanEdit = 777;
            return JsonConvert.SerializeObject(TanEdit);
        }

        //------------CHART---------------------------------------------------------------------------------------------------------------------------------
        public FileContentResult GetChart()
        {
            var dates = new List<Tuple<int, string>>(
                  new[]
                         {
                           new Tuple<int, string> (65, "January"),
                           new Tuple<int, string> (69, "February"),
                           new Tuple<int, string> (90, "March"),
                           new Tuple<int, string> (81, "April"),
                           new Tuple<int, string> (81, "May"),
                           new Tuple<int, string> (55, "June"),
                           new Tuple<int, string> (40, "July")
                         }
                  );

            var chart = new Chart();
            chart.Width = 700;
            chart.Height = 300;
            chart.BackColor = Color.FromArgb(211, 223, 240);
            chart.BorderlineDashStyle = ChartDashStyle.Solid;
            chart.BackSecondaryColor = Color.White;
            chart.BackGradientStyle = GradientStyle.TopBottom;
            chart.BorderlineWidth = 1;
            chart.Palette = ChartColorPalette.BrightPastel;
            chart.BorderlineColor = Color.FromArgb(26, 59, 105);
            chart.RenderType = RenderType.BinaryStreaming;
            chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            chart.AntiAliasing = AntiAliasingStyles.All;
            chart.TextAntiAliasingQuality = TextAntiAliasingQuality.Normal;
            chart.Titles.Add(CreateTitle());
            chart.Legends.Add(CreateLegend());
            chart.Series.Add(CreateSeries(dates, SeriesChartType.Line, Color.Red));
            chart.ChartAreas.Add(CreateChartArea());

            var ms = new MemoryStream();
            chart.SaveImage(ms);
            return File(ms.GetBuffer(), @"image/png");
        }
        Создание заголовка выглядит следующим образом:
[NonAction]
        public Title CreateTitle()
        {
            Title title = new Title();
            title.Text = "Result Chart";
            title.ShadowColor = Color.FromArgb(32, 0, 0, 0);
            title.Font = new Font("Trebuchet MS", 14F, FontStyle.Bold);
            title.ShadowOffset = 3;
            title.ForeColor = Color.FromArgb(26, 59, 105);

            return title;
        }
        Для создания серии воспользуемся методом CreateSeries.
        [NonAction]
public Series CreateSeries(IList<Tuple<int, string>> results,
       SeriesChartType chartType,
       Color color)
        {
            var seriesDetail = new Series();
            seriesDetail.Name = "Result Chart";
            seriesDetail.IsValueShownAsLabel = false;
            seriesDetail.Color = color;
            seriesDetail.ChartType = chartType;
            seriesDetail.BorderWidth = 2;
            seriesDetail["DrawingStyle"] = "Cylinder";
            seriesDetail["PieDrawingStyle"] = "SoftEdge";
            DataPoint point;

            foreach (var result in results)
            {
                point = new DataPoint();
                point.AxisLabel = result.Item2;
                point.YValues = new double[] { result.Item1 };
                seriesDetail.Points.Add(point);
            }
            seriesDetail.ChartArea = "Result Chart";

            return seriesDetail;
        }
        Создание легенды выглядит следующим образом:
[NonAction]
        public Legend CreateLegend()
        {
            var legend = new Legend();
            legend.Name = "Result Chart";
            legend.Docking = Docking.Bottom;
            legend.Alignment = StringAlignment.Center;
            legend.BackColor = Color.Transparent;
            legend.Font = new Font(new FontFamily("Trebuchet MS"), 9);
            legend.LegendStyle = LegendStyle.Row;

            return legend;
        }
        Последним штрихом создадим область, в которой данный график будет отображен.
        [NonAction]
public ChartArea CreateChartArea()
        {
            var chartArea = new ChartArea();
            chartArea.Name = "Result Chart";
            chartArea.BackColor = Color.Transparent;
            chartArea.AxisX.IsLabelAutoFit = false;
            chartArea.AxisY.IsLabelAutoFit = false;
            chartArea.AxisX.LabelStyle.Font = new Font("Verdana,Arial,Helvetica,sans-serif", 8F, FontStyle.Regular);
            chartArea.AxisY.LabelStyle.Font = new Font("Verdana,Arial,Helvetica,sans-serif", 8F, FontStyle.Regular);
            chartArea.AxisY.LineColor = Color.FromArgb(64, 64, 64, 64);
            chartArea.AxisX.LineColor = Color.FromArgb(64, 64, 64, 64);
            chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64);
            chartArea.AxisX.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64);
            chartArea.AxisX.Interval = 1;
            return chartArea;
        }
        //--------------------------------------------------------------------------------------------------------------------------------------------------


    }
}