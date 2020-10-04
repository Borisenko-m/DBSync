using System;
using System.IO;
using Spire.Xls;
using xls = Spire.Xls;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Linq;
using MySql.Data.MySqlClient;
using Excel_Interop.Datasets;
namespace Excel_Interop
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string path = @"c:\Users\davil\OneDrive\Рабочий стол\r.xlsx";
            var connector = new DBConnector();

           


            using (DBSets db = new DBSets())
            {
                
                //добавление регионов
                db.Regions.AddRange(connector.SynchronizeTable<Region>());
                
                Console.WriteLine("1");
                //добавление районов
                db.Districts.AddRange(connector.SynchronizeTable<District>());
                
                Console.WriteLine("2");
                //добавление городов
                db.Cities.AddRange(connector.SynchronizeTable<City>("Cities"));
               
                Console.WriteLine("3");
                //добавление микрорайонов
                db.MicroDistricts.AddRange(connector.SynchronizeTable<MicroDistrict>());
             
                Console.WriteLine("4");
                //добавление улиц
                db.Streets.AddRange(connector.SynchronizeTable<Street>());
              
                Console.WriteLine("5");
                //добавление типов здании
                db.BuildingTypes.AddRange(connector.SynchronizeTable<BuildingType>("Buildings_Types"));
            
                Console.WriteLine("6");
                //добавление здании
                db.Buildings.AddRange(connector.SynchronizeTable<Building>());
        
                Console.WriteLine("7");
                //добавление классификаторов
                db.Classifiers.AddRange(connector.SynchronizeTable<Classifier>());

                Console.WriteLine("8");
                //добавление классификаторов
                db.Users.AddRange(connector.SynchronizeTable<User>());
                Console.WriteLine("9");
                //добавление статусов заявок
                db.ApplicationStatuses.AddRange(connector.SynchronizeTable<ApplicationStatus>("application_statuses"));
                db.SaveChanges();
                Console.WriteLine("10");
                //добавление адреса
                var list = connector.SynchronizeTable<Address>("Addresses");
                db.Addresses.AddRange(list);
              
                Console.WriteLine("11");
                //добавление компании
                db.Companies.AddRange(connector.SynchronizeTable<Company>("managing_companies"));
              

                //добавление компании-классификаторы
                db.CompaniesHasClassifiers.AddRange(connector.SynchronizeTable<CompanyHasClassifier>("join_mc_classifiers"));
          

                //добавление компании-классификаторы
                db.CompaniesHasAddresses.AddRange(connector.SynchronizeTable<CompanyHasAddress>("join_mc_addresses"));
                db.SaveChanges();

                db.Applications.AddRange(connector.SynchronizeTable<Application>());
                db.SaveChanges();




                var rep = new Reports(db);
                var x = rep.GetReport(new DateTime(2019, 12, 1), new DateTime(2020, 12, 12));
                Console.ReadKey();


                //foreach data in List{ new Dataset())
                //db.Dataset.Add();

                //db.Dataset.SaveChanges();
            }

            //var x = new Workbook();
            //x.LoadFromFile(path);
            //var wsh = x.Worksheets[0];
            //var r = wsh.Range["c16"];
            //var arr = wsh.Range["A13:A32"].ToArray();
            //var b = r.HasMerged;



            Console.WriteLine();

            //foreach (var item in arr)
            //{
            //    Console.WriteLine(item.Value);
            //}
            
            //string host = "localhost"; // Имя хоста
            //string database = "eds"; // Имя базы данных
            //string user = "root"; // Имя пользователя
            //string password = "root"; // Пароль пользователя

            //string cnnStr = "Database=" + database + ";Datasource=" + host + ";User=" + user + ";Password=" + password;
            //var mysqlcnn = new MySqlConnection(cnnStr);
            //var query = mysqlcnn.CreateCommand();
            //query.CommandText = "SELECT * FROM users;";
            //mysqlcnn.Open();
            //var res = query.ExecuteReader();
            //int val = 0;
            //while (res.Read())
            //    Console.WriteLine(res.GetFieldValue<string>(1));


            //using (FileStream fs = new FileStream($"{path}user.json", FileMode.OpenOrCreate))
            //{
            //    Col restoredPerson = await JsonSerializer.DeserializeAsync<Col>(fs);
            //    Console.WriteLine(
            //        $"1: {restoredPerson.AddressFact}     \n" +
            //        $"2: {restoredPerson.AddressUr}       \n" +
            //        $"3: {restoredPerson.Director}        \n" +
            //        $"4: {restoredPerson.District}        \n" +
            //        $"5: {restoredPerson.DistrName}       \n" +
            //        $"6: {restoredPerson.Email}           \n" +
            //        $"7: {restoredPerson.ESite}           \n" +
            //        $"8: {restoredPerson.Fax}             \n" +
            //        $"9: {restoredPerson.HCnt}            \n" +
            //        $"10:{restoredPerson.INN}             \n" +
            //        $"11:{restoredPerson.KPP}             \n" +
            //        $"22:{restoredPerson.MO_ID2}          \n" +
            //        $"13:{restoredPerson.mo_id2list}      \n" +
            //        $"14:{restoredPerson.Name}            \n" +
            //        $"15:{restoredPerson.OGRN}            \n" +
            //        $"16:{restoredPerson.PD_ID}           \n" +
            //        $"17:{restoredPerson.PD_Type}         \n" +
            //        $"18:{restoredPerson.PD_LicGive}      \n" +
            //        $"19:{restoredPerson.TelFax}          \n" +
            //        $"20:{restoredPerson.PD_LicNum}       \n");
            //}

            Console.ReadLine();

           // workbook.SaveToFile(@"C:\Users\davil\OneDrive\Рабочий стол\Sample.xlsx", ExcelVersion.Version2016);
        }
        class Col
        {
            public string PD_ID { get; set; }
            public string PD_Type { get; set; }
            public string DistrName { get; set; }
            public string Name { get; set; }
            public string AddressUr { get; set; }
            public string OGRN { get; set; }
            public string TelFax { get; set; }
            public string Email { get; set; }
            public string MO_ID2 { get; set; }
            public string mo_id2list { get; set; }
            public string INN { get; set; }
            public string District { get; set; }
            public string Fax { get; set; }
            public string ESite { get; set; }
            public string PD_LicGive { get; set; }
            public string KPP { get; set; }
            public string AddressFact { get; set; }
            public string Director { get; set; }
            public string PD_LicNum { get; set; }
            public string HCnt { get; set; }
            Col()
            {

            }
        }

        class XlsIO
        {
            Workbook workbook = new Workbook();

            public XlsIO()
            {

            }
            public XlsIO(string workbookName) : this()
            {
                InitWorksheet(workbookName);
            }

            void InitWorksheet(string name)
            {
                foreach (var sheet in workbook.Worksheets)
                {
                    workbook.Worksheets.Remove(sheet);
                }
                workbook.Worksheets.Add(name);
            }

            void Add(Stream stream)
            {

            }

            void Add(string str)
            {

            }

        }
    }
}
