using OSIsoft.AF.Asset;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ilimspecs.Data
{
    public enum VarSpecsChangeStatus { ChangeOK=1, ChangeBad=2, NotChange=3}

    public class Specs : ISpecs
    {
        //Строка соединения
        static string connString;

        //TODO:Изменить обработку ошибок
        static string lastError;
        public static string LastError { get { return lastError; } }

        static Specs()
        {
            try
            {
                connString = ConfigurationManager.ConnectionStrings["specs"].ConnectionString;
            }
            catch (Exception)
            {
            }
            if (connString == null) lastError = "Отсутствует строка соединения к БД";

            //connString = @"Data Source=VMILIM;Initial Catalog=Specs;Integrated Security=True";
        }

        public Specs()
        {
          
        }

        public List<EntityProductionUnit> GetProductUnitAll()
        {
            List<EntityProductionUnit> allProductionUnitList = new List<EntityProductionUnit>();

            using (SqlConnection con = new SqlConnection(connString))
            {
                SqlCommand com = new SqlCommand("SELECT * FROM dbo.ProductionUnits", con);
                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord item in dr)
                        {
                            allProductionUnitList.Add(new EntityProductionUnit { ProductionUnit_Id = (int)item["ProductionUnit_Id"], ProductionUnit_Desc = (string)item["ProductionUnit_Desc"].ToString().ToUpper() });
                        }
                    }
                    else return null;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            return allProductionUnitList;
        }

        public List<EntityProduct> GetAllProductByUnit(string unit)
        {
            List<EntityProduct> lprod = new List<EntityProduct>();

            string query = string.Format("SELECT * FROM Products P INNER JOIN ProductionUnits PU ON P.ProductionUnit_Id=PU.ProductionUnit_Id "+
                                                    "Where PU.ProductionUnit_Desc=@ProductionUnit_Desc AND P.Product_Is_Active=1");
            using (SqlConnection con = new SqlConnection(connString))
            {

                SqlParameter parameter = new SqlParameter()
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = unit,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };

                SqlCommand com = new SqlCommand(query, con);

                com.Parameters.Add(parameter);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord item in dr)
                        {

                            lprod.Add
                                (

                                  //   [Product_Id][bigint] IDENTITY(1, 1) NOT NULL,
                                  //   [Extended_Info] [nvarchar] (255) NULL,
                                  //[Product_Is_Active] [bit] NULL,
                                  //[Product_Code] [nvarchar] (255) NOT NULL,
                                  //    [Product_Desc] [nvarchar] (255) NULL,
                                  //[ProductionUnit_Id] [int] NOT NULL,
                                  //   [Product_SAP] [nchar] (10) NULL,

                                  new EntityProduct
                                  {
                                      Product_Id = (Int64)item["Product_Id"],
                                      Extended_Info = (string)item["Extended_Info"].ToString().ToUpper(),
                                      Product_Is_Active = (bool)item["Product_Is_Active"],
                                      Product_Code = (string)item["Product_Code"].ToString().ToUpper(),
                                      Product_Desc = (string)item["Product_Desc"].ToString().ToUpper(),
                                      ProductionUnit_Id = (int)item["ProductionUnit_Id"],
                                      Product_SAP = (string)item["Product_SAP"].ToString().ToUpper(),
                                      ProductionUnit_Desc = (string)item["ProductionUnit_Desc"].ToString().ToUpper(),
                                      Product_Procont= (string)item["Product_Procont"].ToString()
                                  }
                                );
                        }
                    }
                    else return null;
                }
                catch (Exception e)
                {
                    lastError=e.Message;
                }
            }
            return lprod;

        }

        public List<ListViewItem> GetAllProductByUnitListView(string unit)
        {
            List<ListViewItem> lprodviewitem = new List<ListViewItem>();

            using (SqlConnection con = new SqlConnection(connString))
            {
                string query = string.Format("SELECT * FROM Products P INNER JOIN ProductionUnits PU ON P.ProductionUnit_Id=PU.ProductionUnit_Id " +
                                              "Where PU.ProductionUnit_Desc=@ProductionUnit_Desc AND P.Product_Is_Active=1", unit);

                SqlParameter parameter = new SqlParameter()
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = unit,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };

                SqlCommand com = new SqlCommand(query, con);

                com.Parameters.Add(parameter);
                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord item in dr)
                        {
                            EntityProduct product = new EntityProduct
                            {
                                Product_Id = (Int64)item["Product_Id"],
                                Extended_Info = (string)item["Extended_Info"].ToString().ToUpper(),
                                Product_Is_Active = (bool)item["Product_Is_Active"],
                                Product_Code = (string)item["Product_Code"].ToString().ToUpper(),
                                Product_Desc = (string)item["Product_Desc"].ToString(),
                                ProductionUnit_Id = (int)item["ProductionUnit_Id"],
                                Product_SAP = (string)item["Product_SAP"].ToString(),
                                ProductionUnit_Desc = (string)item["ProductionUnit_Desc"].ToString().ToUpper(),
                                Product_Procont= (string)item["Product_Procont"].ToString()
                            };

                            lprodviewitem.Add(new ListViewItem { Text = product.Product_Code, Tag = product, ToolTipText = product.Product_Desc, Checked=true });
                        }
                    }
                    else return null;
                }
                catch (Exception e)
                {
                    lastError=e.Message;
                }
            }
            return lprodviewitem;
        }

        /// <summary>
        /// Получение всех спецификаций со списком продуктов внутри к которым относятся эти спецификации
        /// </summary>
        /// <param name="lstspecs">Список AFElement-ов, которые являются спецификациями</param>
        /// <returns>Получение списка классов EntitySpecification</returns>
        public List<EntitySpecification> GetAllSpecification(List<EntityVarSpecsAF> lstspecs)
        {
            if (lstspecs == null || lstspecs.Count == 0)
            {
                return null;
            }

            List<EntitySpecification> lstVarSpecs = new List<EntitySpecification>();

            string pitag;
            string pathToItem;

            foreach (EntityVarSpecsAF afitem in lstspecs)
            {
                //TODO: Сделать нормально через константу 
//                pitag = afitem.AFElem.Attributes["Тэг PI"].GetValue().ToString();
                pitag = afitem.AFElem.Attributes[AFEnvironment.PITag].GetValue().ToString();
                //pitag = afitem.Attributes["Тэг PI"].GetValue().ToString();

                EntitySpecification spec = new EntitySpecification();
                spec.Tag = pitag;
                pathToItem = afitem.AFElem.Parent.GetPath();
                spec.PathAF = pathToItem.Substring(pathToItem.IndexOf(AFEnvironment.AFTree));
                spec.Name = afitem.Name;
//                spec.ProductionUnit = afitem.AFElem.Attributes["Код производственного участка"].GetValue().ToString();
                spec.ProductionUnit = afitem.AFElem.Attributes[AFEnvironment.CodeProductionSector].GetValue().ToString();



                string query = string.Format("SELECT * FROM VarSpecs VS INNER JOIN Products P ON VS.Product_Id=P.Product_Id Where (VS.Date_Expiration IS NULL) AND VS.Var_Id = '{0}'", pitag);
                using (SqlConnection con = new SqlConnection(connString))
                {
                    SqlCommand com = new SqlCommand(query, con);
                    try
                    {
                        con.Open();
                        SqlDataReader dr = com.ExecuteReader();

                        if (dr.HasRows)
                        {
                            foreach (DbDataRecord item in dr)
                            {

                                spec.ProductLimits.Add(

                                  new ProductLimits
                                  {
                                      VS_Id = (Int64)item["VS_Id"],
                                      Date_Effective = (DateTime)item["Date_Effective"],
                                      Date_Expiration = (item["Date_Expiration"] == DBNull.Value) ? null : (DateTime?)item["Date_Expiration"],
                                      L_Reject = (item["L_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Reject"].ToString().Replace('.', ',')),
                                      L_Warning = (item["L_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Warning"].ToString().Replace('.', ',')),
                                      Product_Id = (Int64)item["Product_Id"],
                                      Target = (item["Target"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["Target"].ToString().Replace('.', ',')),
                                      U_Reject = (item["U_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Reject"].ToString().Replace('.', ',')),
                                      U_Warning = (item["U_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Warning"].ToString().Replace('.', ',')),
                                      Var_Id = (string)item["Var_Id"].ToString().ToUpper(),
                                      Var_Type = (string)item["Var_Type"].ToString().ToUpper(),
                                      Author = (string)item["Author"].ToString().ToUpper(),


                                      //Product_Id = (Int64)item["Product_Id"],
                                      Extended_Info = (string)item["Extended_Info"].ToString().ToUpper(),
                                      Product_Is_Active = (bool)item["Product_Is_Active"],
                                      Product_Code = (string)item["Product_Code"].ToString().ToUpper(),
                                      Product_Desc = (string)item["Product_Desc"].ToString().ToUpper(),
                                      ProductionUnit_Id = (int)item["ProductionUnit_Id"],
                                      Product_SAP = (string)item["Product_SAP"].ToString().ToUpper(),
                                      Product_Procont=(string)item["Product_Procont"].ToString()
                                  }
                                );
                            }
                         }
                        //else return null; //убрал т.к. если в Specs по Var_Id отсутствуют границы получается выход
                    }
                    
                    catch (Exception e)
                    {
                        lastError= e.Message;
                    }
                    
                }
                lstVarSpecs.Add(spec);
            }
            return lstVarSpecs;
        }

        /// <summary>
        /// Создание нового продукта в табл Products в БД Specs
        /// </summary>
        /// <param name="Product_Code">Код обновляемого продукта</param>
        /// <param name="Product_Desc">Описание обновляемого продукта</param>
        /// <param name="ProductionUnit_Desc">Название оборудования на котором производится продукт</param>
        /// <returns></returns>
        public int SaveNewProduct(string Product_Code, string Product_Desc, string ProductionUnit_Desc, string product_SAP, string product_Procont)
        {
            // На таблицу Products добавлено ограничение уникальности
            // alter table [dbo].[Products]
            // add constraint UK_productID_unitID UNIQUE([Product_Code], [ProductionUnit_Id])

            int Result = 0;
            using (SqlConnection con = new SqlConnection(connString))
            {
                //строка запроса через параметры
                string query = string.Format("INSERT dbo.Products(Product_Is_Active, Product_Code, Product_Desc, Product_SAP , Product_Procont, ProductionUnit_Id) " +
                   "SELECT 1, @Product_Code, @Product_Desc, @Product_SAP, @Product_Procont, (SELECT TOP 1 ProductionUnit_Id FROM dbo.ProductionUnits WHERE ProductionUnit_Desc=@ProductionUnit_Desc); " +
                   "SELECT CAST(@@IDENTITY as int) as ID ");

                //INSERT[Specs].[dbo].[Products]([Product_Is_Active], [Product_Code], [Product_Desc], [ProductionUnit_Id])
                //SELECT 1, @Product_Code, @Product_Desc, (SELECT TOP 1 [ProductionUnit_Id] FROM[Specs].[dbo].[ProductionUnits] WHERE[ProductionUnit_Desc]=@ProductionUnit_Desc)

                SqlCommand com = new SqlCommand(query, con);

                object obj;

                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@Product_Code",
                    Value = Product_Code,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_Desc",
                    Value = Product_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_SAP",
                    Value = product_SAP,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_Procont",
                    Value = product_Procont,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();

                    obj = com.ExecuteScalar();
                    if (obj != null)
                    {
                        Result = (int)obj;
                    }
                    else
                    {
                        Result = 0;
                    }


                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }

            }
            return Result;
        }

        /// <summary>
        /// Обновление продуктов в таблице Products БД Specs
        /// </summary>
        /// <param name="Product_Id">ID обновляемого продукта</param>
        /// <param name="Product_Code">Код обновляемого продукта</param>
        /// <param name="Product_Desc">Описание обновляемого продукта</param>
        /// <param name="ProductionUnit_Desc">Название оборудования на котором производится продукт</param>
        /// <param name="product_SAP">Имя продукта в SAP</param>
        /// <param name="product_Procont">Имя продукта в Procont</param>
        /// <returns></returns>
        // (product_Id, product_Code.Trim(), product_Desc.Trim(), productionUnit_Desc.Trim(), product_SAP.Trim(), product_Procont.Trim())
        public int UpDateProduct(int Product_Id, string Product_Code, string Product_Desc, string ProductionUnit_Desc, string product_SAP,string product_Procont)
        {
            // На таблицу Products добавлено ограничение уникальности
            // alter table [dbo].[Products]
            // add constraint UK_productID_unitID UNIQUE([Product_Code], [ProductionUnit_Id])

            int flagResult = 0;
            using (SqlConnection con = new SqlConnection(connString))
            {
                //строка запроса через параметры
                string query = string.Format("UPDATE P SET P.[Product_Desc]=@Product_Desc, P.[Product_Code]=@Product_Code, P.[Product_SAP]=@Product_SAP, P.[Product_Procont]=@Product_Procont " +
                   "FROM dbo.Products P " +
                   "INNER JOIN Specs.dbo.ProductionUnits PU on P.ProductionUnit_Id= PU.ProductionUnit_Id " +
                   "Where Product_Id =@Product_Id AND PU.ProductionUnit_Desc=@ProductionUnit_Desc ");

                //UPDATE P SET P.[Product_Desc]= 'ddddd', P.[Product_Code]='Glossy 90 gsm 22'
                //FROM[Specs].[dbo].[Products] P
                //INNER JOIN[Specs].[dbo].[ProductionUnits] PU on P.ProductionUnit_Id= PU.ProductionUnit_Id
                //Where[Product_Id] = 18 AND PU.[ProductionUnit_Desc]= 'OMC'

                SqlCommand com = new SqlCommand(query, con);

                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@Product_Id",
                    Value = Product_Id,
                    SqlDbType = System.Data.SqlDbType.Int
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_Code",
                    Value = Product_Code,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_Desc",
                    Value = Product_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);


                param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_SAP",
                    Value = product_SAP,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@Product_Procont",
                    Value = product_Procont,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);



                try
                {
                    con.Open();
                    if (com.ExecuteNonQuery() == 1) flagResult = 1;
                    else flagResult = 0;
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }
            }
            return flagResult;
        }

        /// <summary>
        /// Возвращает ProductID из табл Product
        /// </summary>
        /// <param name="Product_Code">Код подукта для которого ищется ProductID</param>
        /// <param name="ProductionUnit_Desc">Название оборудования на котором производится продукт</param>
        /// <returns></returns>
        public Int64 GetProductID(string Product_Code, string ProductionUnit_Desc)
        {
            Int64 ResultProductionID = 0;



            //select P.Product_Id
            //from[dbo].[Products] P
            //inner join[dbo].[ProductionUnits] PU ON P.ProductionUnit_Id=PU.ProductionUnit_Id
            //where P.Product_Code='GLOSSY 115 GSM' AND PU.ProductionUnit_Desc='OMC'

            using (SqlConnection con = new SqlConnection(connString))
            {
                //строка запроса через параметры
                string query = string.Format("SELECT TOP 1 P.[Product_Id] AS Product_Id FROM dbo.Products P " +
                                                  "INNER JOIN dbo.ProductionUnits PU ON P.[ProductionUnit_Id]=PU.[ProductionUnit_Id] " +
                                                  "WHERE P.[Product_Code]=@Product_Code AND PU.[ProductionUnit_Desc]=@ProductionUnit_Desc");

                SqlCommand com = new SqlCommand(query, con);


                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@Product_Code",
                    Value = Product_Code.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord item in dr)
                        {
                            ResultProductionID = (Int64)item["Product_Id"];
                            break;
                        }
                    }
                    else
                    {
                        //Если записи из БД не вернулись, то Var_Id значит по этому Var_Id нет границ по продуктам
                    }
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }
            }
            //Закончилось первое обращение к БД Specs относительно страрой спецификации

            return ResultProductionID;
        }

        /// <summary>
        /// Метод для сравнения двух границ. Используется для сравнения границ с листа с границей в БД, возвращает false если границы не равны
        /// </summary>
        /// <param name="newProductLimit">Граница с листа Excel</param>
        /// <param name="oldProductLimit">Граница с БД</param>
        private bool CompareProductLimit(ProductLimit newProductLimit, ProductLimit oldProductLimit)
        {
            bool ResultFlag = true;

            if (newProductLimit.LoLoRule)
            {
                if (newProductLimit.LoLo!= oldProductLimit.LoLo)
                {
                    ResultFlag = false;
                    return ResultFlag;
                }
                else
                {
                    ResultFlag = true;
                }
            }

            if (newProductLimit.LoRule)
            {
                if (newProductLimit.Lo != oldProductLimit.Lo)
                {
                    ResultFlag = false;
                    return ResultFlag;
                }
                else
                {
                    ResultFlag = true;
                }
            }

            if (newProductLimit.TargetRule)
            {
                if (newProductLimit.Target != oldProductLimit.Target)
                {
                    ResultFlag = false;
                    return ResultFlag;
                }
                else
                {
                    ResultFlag = true;
                }
            }

            if (newProductLimit.HiRule)
            {
                if (newProductLimit.Hi != oldProductLimit.Hi)
                {
                    ResultFlag = false;
                    return ResultFlag;
                }
                else
                {
                    ResultFlag = true;
                }
            }

            if (newProductLimit.HiHiRule)
            {
                if (newProductLimit.HiHi != oldProductLimit.HiHi)
                {
                    ResultFlag = false;
                    return ResultFlag;
                }
                else
                {
                    ResultFlag = true;
                }
            }

            return ResultFlag;
        }


        /// <summary>
        /// Обновление или создание(если их нет) границ табл VarSpecs в БД Specs по странной логике
        /// </summary>
        /// <param name="prlimit">ProductLimit класс который содержит в себе какие границы, по какому оборудованию и для какого тега нужна изменить</param>
        /// <returns></returns>
        public VarSpecsChangeStatus UpDateVarSpecs(ProductLimit prlimit)
        {
            VarSpecsChangeStatus flagResult = VarSpecsChangeStatus.NotChange;
            ProductLimit productLimitOld = null;
            Int64 VS_Id = 0; //Переменная для хранения ID спецификации которую изменяем
            Int64 Product_Id = 0; //Переменная для хранения ID продукта

            decimal? valueToWrite;//Переменная для записи значения в БД


            using (SqlConnection con = new SqlConnection(connString))
            {
                //строка запроса через параметры
                //получение границ через имя тега, имя производящего оборудования, кода продукта
                string query = string.Format("SELECT * FROM VarSpecs VS INNER JOIN Products P ON VS.Product_Id=P.Product_Id " +
                                                  "INNER JOIN ProductionUnits PU ON P.ProductionUnit_Id=PU.ProductionUnit_Id " +
                               "Where (VS.Date_Expiration IS NULL) AND P.Product_Code=@Product_Code AND PU.ProductionUnit_Desc=@ProductionUnit_Desc  AND VS.Var_Id=@Var_Id");

                SqlCommand com = new SqlCommand(query, con);


                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@Product_Code",
                    Value = prlimit.ProductCode.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = prlimit.ProductionUnit_Desc.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);


                param = new SqlParameter
                {
                    ParameterName = "@Var_Id",
                    Value = prlimit.Tag.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    productLimitOld = new ProductLimit(prlimit.ProductionUnit_Desc, prlimit.Tag, prlimit.ProductCode, prlimit.Author);

                    if (dr.HasRows)
                    {
                        //Кто был старым автором границы нас не интересует

                        foreach (DbDataRecord item in dr)
                        {
                            VS_Id = (Int64)item["VS_Id"];
                            Product_Id = (Int64)item["Product_Id"];
                            productLimitOld.LoLo = (item["L_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Reject"].ToString().Replace('.', ','));
                            productLimitOld.Lo = (item["L_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Warning"].ToString().Replace('.', ','));
                            productLimitOld.Target = (item["Target"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["Target"].ToString().Replace('.', ','));
                            productLimitOld.Hi = (item["U_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Warning"].ToString().Replace('.', ','));
                            productLimitOld.HiHi = (item["U_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Reject"].ToString().Replace('.', ','));
                            productLimitOld.Author = (string)item["Author"].ToString().ToUpper();

                            if (CompareProductLimit(prlimit, productLimitOld))
                            {
                                flagResult = VarSpecsChangeStatus.NotChange;
                                return flagResult;
                            }
                        }

                    }
                    else
                    {
                        //Если записи из БД не вернулись, по Var_Id значит по этому Var_Id нет границ по продуктам
                        //Тогда надо получить ProductID и создать новые границы по этому продукту
                        Int64 tempProductID = GetProductID(prlimit.ProductCode.Trim(), prlimit.ProductionUnit_Desc.Trim());

                        if (tempProductID!=0)
                        {
                            VS_Id = 0;
                            Product_Id = tempProductID;
                            productLimitOld.LoLo = null;
                            productLimitOld.Lo = null;
                            productLimitOld.Target = null;
                            productLimitOld.Hi = null;
                            productLimitOld.HiHi = null;
                            productLimitOld.Author = null;
                        }
                        else
                        {
                            //То значит беда значит беда нет ProductID по этому коду продукта(ProductCode) и имени оборудования(ProductionUnit_Desc)
                            //и выходим с ошибкой
                            lastError = "Для донного кода продукта (" + prlimit.ProductCode.Trim() + ") и оборудования (" + prlimit.ProductionUnit_Desc.Trim() + "), нет ProductID в БД Specs";
                            return flagResult=0;
                        }

                        //TODO: Если записи не вернулись надо определить Product_ID для записи в VarSpecs по ниже указанному запросу
                        //select P.Product_Id
                        //from[dbo].[Products] P
                        //inner join[dbo].[ProductionUnits] PU ON P.ProductionUnit_Id=PU.ProductionUnit_Id
                        //where P.Product_Code='GLOSSY 115 GSM' AND PU.ProductionUnit_Desc='OMC'
                    }
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }
                finally
                {

                }
            }
            //Закончилось первое обращение к БД Specs относительно страрой спецификации

            //Обновляем старые границы чтобы они стали новыми границами
            if (productLimitOld != null)
            {
                if (prlimit.HiHiRule) productLimitOld.HiHi = prlimit.HiHi;
                if (prlimit.HiRule) productLimitOld.Hi = prlimit.Hi;
                if (prlimit.TargetRule) productLimitOld.Target = prlimit.Target;
                if (prlimit.LoRule) productLimitOld.Lo = prlimit.Lo;
                if (prlimit.LoLoRule) productLimitOld.LoLo = prlimit.LoLo;
                productLimitOld.Author = prlimit.Author;
            }

            if (!productLimitOld.CheckLimits())
            {
                //TODO:Переделать обработку ошибок при сравнении границ
                //Устанавливаем ошибку и выходим из метода
                lastError = "Сравнение границ привело к ошибке";
                return flagResult=VarSpecsChangeStatus.ChangeBad;
            }

            using (SqlConnection con = new SqlConnection(connString))
            {

                //Начинаю Транзанкцию
                SqlCommand cmdUpdate = new SqlCommand(String.Format("UPDATE VarSpecs SET Date_Expiration=GETDATE() WHERE VS_Id=@VS_Id"), con);
                SqlCommand cmdInsert = new SqlCommand(string.Format("INSERT INTO VarSpecs (Date_Effective, L_Reject, L_Warning, Target,U_Warning, U_Reject,Var_Id, Author, Product_Id  )" +
                    " VALUES (GETDATE(), @L_Reject, @L_Warning, @Target, @U_Warning, @U_Reject, @Var_Id, @Author, @Product_Id)"), con);

                //if (productLimitOld == null) valueToWrite = DBNull.Value;
                //else valueToWrite = productLimitOld.LoLo,

                SqlParameter paramTran = new SqlParameter
                {
                    ParameterName = "@VS_Id",
                    Value = VS_Id,
                    SqlDbType = System.Data.SqlDbType.BigInt
                };
                cmdUpdate.Parameters.Add(paramTran);


                paramTran = new SqlParameter();
                paramTran.ParameterName = "@L_Reject";
                if (productLimitOld.LoLo == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = productLimitOld.LoLo; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@L_Warning";
                if (productLimitOld.Lo == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = productLimitOld.Lo; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);


                paramTran = new SqlParameter();
                paramTran.ParameterName = "@Target";
                if (productLimitOld.Target == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = productLimitOld.Target; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@U_Warning";
                if (productLimitOld.Hi == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = productLimitOld.Hi; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@U_Reject";
                if (productLimitOld.HiHi == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = productLimitOld.HiHi; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                //@Var_Id
                paramTran = new SqlParameter
                {
                    ParameterName = "@Var_Id",
                    Value = productLimitOld.Tag,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@Author
                paramTran = new SqlParameter
                {
                    ParameterName = "@Author",
                    Value = productLimitOld.Author,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@Product_Id
                paramTran = new SqlParameter
                {
                    ParameterName = "@Product_Id",
                    Value = Product_Id,
                    SqlDbType = System.Data.SqlDbType.BigInt
                };
                cmdInsert.Parameters.Add(paramTran);


                // Получаем из объекта подключения.
                SqlTransaction tx = null;
                try
                {
                    con.Open();
                    tx = con.BeginTransaction();
                    // Включение команд в транзакцию
                    cmdUpdate.Transaction = tx;
                    cmdInsert.Transaction = tx;
                    // Выполнение команд.

                    //Если границы (VS_Id!=0) существую то мы их обновляем
                    if (VS_Id!=0) cmdUpdate.ExecuteNonQuery();
                    if (cmdInsert.ExecuteNonQuery()==1)
                    {
                        flagResult = VarSpecsChangeStatus.ChangeOK;
                    }
                    else
                    {
                        flagResult = VarSpecsChangeStatus.ChangeBad;
                    }
                    // Фиксация.
                    tx.Commit();
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                    // При возникновении любой ошибки выполняется откат транзакции.
                    tx.Rollback();
                }
            }
            return flagResult;
        }

        /// <summary>
        /// Сохранение в БД новой единицы оборудования (ProductionUnit) 
        /// </summary>
        /// <param name="ProductionUnit_Desc">Имя оборудования сохраняемого в БД Specs</param>
        /// <returns>Возвращает ID записанной ProductionUnit</returns>
        public int SaveNewProductionUnitAndGet(string ProductionUnit_Desc)
        {
            // Перед сохранением нужно проверить что такой ProductionUnit есть в AF
            // Добавить ограничение уникальности для ProductionUnit_Desc

            int Result = 0;

            using (SqlConnection con = new SqlConnection(connString))
            {
                //строка запроса через параметры 
                //string query = string.Format("INSERT INTO dbo.ProductionUnits (ProductionUnit_Desc) VALUES (@ProductionUnit_Desc); SELECT SCOPE_IDENTITY();");
                //string query = string.Format("INSERT INTO dbo.ProductionUnits (ProductionUnit_Desc) VALUES (@ProductionUnit_Desc); SELECT @@IDENTITY as ID;");

                string query = string.Format("if ((select ProductionUnit_Id from dbo.ProductionUnits where ProductionUnit_Desc = @ProductionUnit_Desc) Is Null) " +
                    "begin " +
                    "INSERT INTO dbo.ProductionUnits(ProductionUnit_Desc) VALUES(@ProductionUnit_Desc); SELECT CAST(@@IDENTITY as int) as ID " +
                    "end " +
                    "else " +
                    "Select ProductionUnit_Id from dbo.ProductionUnits where ProductionUnit_Desc = @ProductionUnit_Desc");

                object obj;

                SqlCommand com = new SqlCommand(query, con);

                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();
                    obj = com.ExecuteScalar();
                    if (obj != null)
                    {
                        Result = (int)obj;
                    }
                    else
                    {
                        Result = 0;
                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.Message);
                }

            }
            return Result;
        }

        /// <summary>
        /// Обновление в БД единицы оборудования (ProductionUnit)
        /// </summary>
        /// <param name="ProductionUnit_Id">Код оборудования которое нужно обновить</param>
        /// <param name="ProductionUnit_Desc">Новое имя оборудования сохраняемого в БД Specs</param>
        /// <returns></returns>
        public int UpDateProductionUnit(int ProductionUnit_Id, string ProductionUnit_Desc)
        {
            int flagResult = 0;
            using (SqlConnection con = new SqlConnection(connString))
            {
                //строка запроса через параметры
                string query = string.Format("UPDATE dbo.ProductionUnits SET ProductionUnit_Desc=@ProductionUnit_Desc Where ProductionUnit_Id=@ProductionUnit_Id");

                SqlCommand com = new SqlCommand(query, con);

                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Id",
                    Value = ProductionUnit_Id,
                    SqlDbType = System.Data.SqlDbType.Int
                };
                com.Parameters.Add(param);

                param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();

                    //Получение ID последней вставленной записи
                    //"INSERT INTO Model (title, author, line, length, step, date_create) VALUES (@Title, @Author, @Line, @Length, @Step, @Date_create)" SELECT @@IDENTITY AS Id;

                    if (com.ExecuteNonQuery() == 1) flagResult = 1;
                }
                catch (Exception e)
                {
                    lastError= e.Message;
                }
            }
            return flagResult;
        }

        public int GetCountOfProductByProductionUnitDesc(string ProductionUnit_Desc)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Возвращает список спецификаций для которых все границы одинаковы
        /// </summary>
        /// <param name="ProductionUnit_Desc">Название оборудования для кот ищутся одинаковые границы для продуктов производимых на нем</param>
        /// <returns></returns>
        public List<EntityVarSpecCommonBound> GetCommonBound(string ProductionUnit_Desc)
        {
            List<EntityVarSpecCommonBound> listCommonBound = new List<EntityVarSpecCommonBound>();

            if (ProductionUnit_Desc == null )
            {
                return null;
            }

            string query = string.Format("WITH " +
                     "VarSpec_Var_ID(Var_Id) AS " +
                    "(SELECT VS.Var_Id AS Var_Id FROM [dbo].[VarSpecs] VS " +
                     "LEFT JOIN [dbo].[Products] P ON P.Product_Id = VS.Product_Id " +
                     "LEFT JOIN [dbo].[ProductionUnits] PU ON PU.ProductionUnit_Id = P.ProductionUnit_Id " +
                     "WHERE PU.ProductionUnit_Desc = @ProductionUnit_Desc AND VS.[Date_Expiration] IS NULL " +
                     "GROUP BY VS.Var_Id), " +
                    "VarSpec_Var_ID_with_Limits(Var_Id, L_Reject, L_Warning, [Target], [U_Reject], [U_Warning], CountProduct) AS " +
                     "(SELECT VS.Var_Id AS Var_Id, [L_Reject], [L_Warning], [Target],[U_Reject], [U_Warning], Count(VS.Product_Id) AS CountProduct FROM [dbo].[VarSpecs] VS " +
                     "LEFT JOIN [dbo].[Products] P ON P.Product_Id = VS.Product_Id " +
                     "LEFT JOIN [dbo].[ProductionUnits] PU ON PU.ProductionUnit_Id = P.ProductionUnit_Id " +
                     "WHERE PU.ProductionUnit_Desc = @ProductionUnit_Desc AND VS.[Date_Expiration] IS NULL " +
                     "GROUP BY VS.Var_Id, [Var_Id], [L_Reject], [L_Warning], [Target],[U_Reject], [U_Warning] " +
                     "HAVING  Count(VS.Product_Id) = (SELECT count([Product_Id]) " +
                      "FROM [dbo].[Products] " +
                      "WHERE ProductionUnit_Id = (Select TOP 1 ProductionUnit_Id from[dbo].[ProductionUnits] WHERE ProductionUnit_Desc = @ProductionUnit_Desc))) " +
                    "SELECT VI.Var_Id, VIL.L_Reject, VIL.L_Warning, VIL.[Target], VIL.U_Warning, VIL.U_Reject, VIL.CountProduct " +
                    "FROM VarSpec_Var_ID AS VI " +
                    "LEFT JOIN VarSpec_Var_ID_with_Limits AS VIL " +
                    "ON VI.Var_Id = VIL.Var_Id " +
                    "WHERE CountProduct IS NOT NULL");

                using (SqlConnection con = new SqlConnection(connString))
                {
                SqlCommand com = new SqlCommand(query, con);

                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                        con.Open();
                        SqlDataReader dr = com.ExecuteReader();

                        if (dr.HasRows)
                        {
                            foreach (DbDataRecord item in dr)
                            {

                            listCommonBound.Add
                                (
                                  new EntityVarSpecCommonBound
                                  {
                                      Var_Id = (string)item["Var_Id"].ToString().ToUpper(),
                                      L_Reject = (item["L_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Reject"].ToString().Replace('.', ',')),
                                      L_Warning = (item["L_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Warning"].ToString().Replace('.', ',')),
                                      Target = (item["Target"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["Target"].ToString().Replace('.', ',')),
                                      U_Reject = (item["U_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Reject"].ToString().Replace('.', ',')),
                                      U_Warning = (item["U_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Warning"].ToString().Replace('.', ',')),
                                      CountProduct = (item["CountProduct"] == DBNull.Value) ? null : (int?)Convert.ToInt32(item["CountProduct"])
                                  }
                                );
                            }
                        }
                        else return null; //Если запрос не вернул общие записи, значит общие границы отсутствуют
                    }
                    catch (Exception e)
                    {
                        lastError = e.Message;
                    }
                }
            return listCommonBound;
        }

        /// <summary>
        /// Возвращает список спецификаций для которых все границы одинаковы
        /// </summary>
        /// <param name="ProductionUnit_Desc">Название оборудования для кот ищутся одинаковые границы для продуктов производимых на нем</param>
        /// <param name="checkLimits">Границы по которым будут искаться</param>
        /// <returns></returns>
        public List<EntityVarSpecCommonBound> GetCommonBoundCheckLimits(string ProductionUnit_Desc, CheckLimits checkLimits)
        {
            List<EntityVarSpecCommonBound> listCommonBound = new List<EntityVarSpecCommonBound>();


            if (ProductionUnit_Desc == null)
            {
                return null;
            }

            string strquiery1 = string.Empty;
            string strquiery2 = string.Empty;

            if (checkLimits.ShowLoLo)
            {
                strquiery1 += ", L_Reject ";
                strquiery2 += ", VIL.L_Reject ";
            }
            if (checkLimits.ShowLo)
            {
                strquiery1 += ",L_Warning ";
                strquiery2 += ", VIL.L_Warning ";
            }
            if (checkLimits.ShowTarget)
            {
                strquiery1 += ",[Target] ";
                strquiery2 += ", VIL.[Target] ";
            }
            if (checkLimits.ShowHiHi)
            {
                strquiery1 += ",[U_Reject] ";
                strquiery2 += ", VIL.U_Reject ";
            }
            if (checkLimits.ShowHi)
            {
                strquiery1 += ",[U_Warning] ";
                strquiery2 += ", VIL.U_Warning ";
            }

            string query = string.Format("WITH " +
                     "VarSpec_Var_ID(Var_Id) AS " +
                    "(SELECT VS.Var_Id AS Var_Id FROM [dbo].[VarSpecs] VS " +
                     "LEFT JOIN [dbo].[Products] P ON P.Product_Id = VS.Product_Id " +
                     "LEFT JOIN [dbo].[ProductionUnits] PU ON PU.ProductionUnit_Id = P.ProductionUnit_Id " +
                     "WHERE PU.ProductionUnit_Desc = @ProductionUnit_Desc AND VS.[Date_Expiration] IS NULL " +
                     "GROUP BY VS.Var_Id), " +
                    "VarSpec_Var_ID_with_Limits(Var_Id "+ strquiery1+", CountProduct) AS " +
                     "(SELECT VS.Var_Id AS Var_Id "+ strquiery1 + ", Count(VS.Product_Id) AS CountProduct FROM [dbo].[VarSpecs] VS " +
                     "LEFT JOIN [dbo].[Products] P ON P.Product_Id = VS.Product_Id " +
                     "LEFT JOIN [dbo].[ProductionUnits] PU ON PU.ProductionUnit_Id = P.ProductionUnit_Id " +
                     "WHERE PU.ProductionUnit_Desc = @ProductionUnit_Desc AND VS.[Date_Expiration] IS NULL " +
                     "GROUP BY VS.Var_Id, [Var_Id] " + strquiery1 +
                     "HAVING  Count(VS.Product_Id) = (SELECT count([Product_Id]) " +
                      "FROM [dbo].[Products] " +
                      "WHERE ProductionUnit_Id = (Select TOP 1 ProductionUnit_Id from[dbo].[ProductionUnits] WHERE ProductionUnit_Desc = @ProductionUnit_Desc))) " +
                    "SELECT VI.Var_Id "+ strquiery2 + ", VIL.CountProduct " + 
                    "FROM VarSpec_Var_ID AS VI " +
                    "LEFT JOIN VarSpec_Var_ID_with_Limits AS VIL " +
                    "ON VI.Var_Id = VIL.Var_Id " +
                    "WHERE CountProduct IS NOT NULL");

            using (SqlConnection con = new SqlConnection(connString))
            {
                SqlCommand com = new SqlCommand(query, con);

                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord item in dr)
                        {
                            EntityVarSpecCommonBound entityVarSpecCommonBound=   new EntityVarSpecCommonBound();
                            
                            entityVarSpecCommonBound.Var_Id = (string)item["Var_Id"].ToString().ToUpper();
                            entityVarSpecCommonBound.CountProduct = (item["CountProduct"] == DBNull.Value) ? null : (int?)Convert.ToInt32(item["CountProduct"]);

                            if (checkLimits.ShowLoLo)
                            {
                                entityVarSpecCommonBound.L_Reject = (item["L_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Reject"].ToString().Replace('.', ','));
                            }
                            else
                            {
                                entityVarSpecCommonBound.L_Reject = null;
                            }
                            if (checkLimits.ShowLo)
                            {
                                entityVarSpecCommonBound.L_Warning = (item["L_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Warning"].ToString().Replace('.', ','));
                            }
                            else
                            {
                                entityVarSpecCommonBound.L_Warning = null;
                            }
                            if (checkLimits.ShowTarget)
                            {
                                entityVarSpecCommonBound.Target = (item["Target"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["Target"].ToString().Replace('.', ','));
                            }
                            else
                            {
                                entityVarSpecCommonBound.Target = null;
                            }
                            if (checkLimits.ShowHiHi)
                            {
                                entityVarSpecCommonBound.U_Reject = (item["U_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Reject"].ToString().Replace('.', ','));
                            }
                            else
                            {
                                entityVarSpecCommonBound.U_Reject = null;
                            }
                            if (checkLimits.ShowHi)
                            {
                                entityVarSpecCommonBound.U_Warning = (item["U_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Warning"].ToString().Replace('.', ','));
                            }
                            else
                            {
                                entityVarSpecCommonBound.U_Warning = null;
                            }

                            listCommonBound.Add(entityVarSpecCommonBound);

                            //{
                            //    
                            //    L_Reject = (item["L_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Reject"].ToString().Replace('.', ',')),
                            //    L_Warning = (item["L_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Warning"].ToString().Replace('.', ',')),
                            //    Target = (item["Target"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["Target"].ToString().Replace('.', ',')),
                            //    U_Reject = (item["U_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Reject"].ToString().Replace('.', ',')),
                            //    U_Warning = (item["U_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Warning"].ToString().Replace('.', ',')),
                            //    
                            //}

                        }
                    }
                    else return null; //Если запрос не вернул общие записи, значит общие границы отсутствуют
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }
            }
            return listCommonBound;
        }

        /// <summary>
        /// Получение списка границ по продукту по имени тега и имени производящего оборудования
        /// </summary>
        /// <param name="Tag">Имя тега</param>
        /// <param name="ProductionUnit_Desc">Имя производящего оборудования</param>
        /// <returns></returns>
        public List<ProductLimitCommondBound> GetProductLimitCommondBound(string Tag, string ProductionUnit_Desc)
        {
            List<ProductLimitCommondBound> lstProductLimitCommondBound = new List<ProductLimitCommondBound>();

            using (SqlConnection con = new SqlConnection(connString))
            {

                string query = string.Format("SELECT * FROM VarSpecs VS " +
                                             "INNER JOIN Products P ON VS.Product_Id = P.Product_Id " +
                                             "INNER JOIN ProductionUnits PU ON P.ProductionUnit_Id = PU.ProductionUnit_Id " +
                                             "Where(VS.Date_Expiration IS NULL) AND PU.ProductionUnit_Desc = @ProductionUnit_Desc  AND VS.Var_Id = @Tag");

                SqlCommand com = new SqlCommand(query, con);


                SqlParameter param = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = ProductionUnit_Desc.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);


                param = new SqlParameter
                {
                    ParameterName = "@Tag",
                    Value = Tag.Trim(),
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                com.Parameters.Add(param);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader();

                    if (dr.HasRows)
                    {
                        //Кто был старым автором границы нас не интересует

                        foreach (DbDataRecord item in dr)
                        {
                            ProductLimitCommondBound productLimitCommondBound = new ProductLimitCommondBound(ProductionUnit_Desc,Tag,null,null);
                            productLimitCommondBound.VS_Id = (Int64)item["VS_Id"];
                            productLimitCommondBound.LoLo= (item["L_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Reject"].ToString().Replace('.', ','));
                            productLimitCommondBound.Lo = (item["L_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["L_Warning"].ToString().Replace('.', ','));
                            productLimitCommondBound.Target= (item["Target"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["Target"].ToString().Replace('.', ','));
                            productLimitCommondBound.Hi= (item["U_Warning"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Warning"].ToString().Replace('.', ','));
                            productLimitCommondBound.HiHi = (item["U_Reject"] == DBNull.Value) ? null : (decimal?)Convert.ToDouble(item["U_Reject"].ToString().Replace('.', ','));
                            productLimitCommondBound.Author= (string)item["Author"].ToString();
                            productLimitCommondBound.Product_Id= (Int64)item["Product_Id"];
                            productLimitCommondBound.ProductCode = (string)item["Product_Code"];

                            lstProductLimitCommondBound.Add(productLimitCommondBound);
                        }
                    }
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }

            }
            return lstProductLimitCommondBound;
        }

        /// <summary>
        /// Обновление границ для всех продуктов для выбранного тега на определенном оборудовании
        /// </summary>
        /// <param name="prlimit">Класс ProductLimit который содежит в себе имя тега и имя оборудования</param>
        /// <returns></returns>
        public VarSpecsChangeStatus UpDateVarSpecsCommonBound(ProductLimit prlimit)
        {
            VarSpecsChangeStatus flagResult = VarSpecsChangeStatus.NotChange;

            //Получаю продукты с границами, которые установлены в VarSpecs по определенному тегу
            List<ProductLimitCommondBound> lstCommonBoundForTag = GetProductLimitCommondBound(prlimit.Tag, prlimit.ProductionUnit_Desc);

            //Получаю список всех продуктов по которым должны быть границы
            List<EntityProduct> lstProducts= GetAllProductByUnit(prlimit.ProductionUnit_Desc);

            int resultCreateUdateToDBp;

            ProductLimit productLimitForWriteSpecs;

            foreach (EntityProduct item in lstProducts)
            {
                ProductLimitCommondBound productLimitCommondBound = lstCommonBoundForTag.Find(x => String.Compare(x.ProductCode, item.Product_Code, true) == 0);
            
                if (productLimitCommondBound != null)
                {
                    //Если productLimitCommondBound НЕ Null значит границы по этому продукту были и их надо обновить если они разные

                    //сравниваем границы по тегу и продукту с данными в БД и если они разные запускаем обновление
                    if (!CompareProductLimit(prlimit, (ProductLimit)productLimitCommondBound))
                    {
                        productLimitForWriteSpecs = ProductLimitToWrtiteSpecs(prlimit, productLimitCommondBound);

                        //Проверяем валидность границ если все ОК то записываем
                        if (productLimitForWriteSpecs.CheckLimits())
                        {
                            resultCreateUdateToDBp = UpdateVarSpecsOneBound(productLimitForWriteSpecs, productLimitCommondBound.VS_Id, productLimitCommondBound.Product_Id);
                            if (resultCreateUdateToDBp == 1)
                            {
                                flagResult = VarSpecsChangeStatus.ChangeOK;
                            }
                            else
                            {
                                flagResult = VarSpecsChangeStatus.ChangeBad;
                            }
                        }
                        else
                        {
                            flagResult = VarSpecsChangeStatus.ChangeBad;
                        }
                    }
                    else
                    {
                        flagResult = VarSpecsChangeStatus.NotChange;
                    }
                }
                else
                {
                   
                    //Если entityProduct = NULL значит границ нету и их надо создать
                    resultCreateUdateToDBp = CreateVarSpecsOneBound(prlimit, item);
                    if (resultCreateUdateToDBp == 1)
                    {
                        flagResult = VarSpecsChangeStatus.ChangeOK;
                    }
                    else
                    {
                        flagResult = VarSpecsChangeStatus.ChangeBad;
                    }
                }

            }
            return flagResult;
        }

        /// <summary>
        /// Метод формирования границ из старых и новых границ, используется в общих границах
        /// </summary>
        /// <param name="prlimit">Границы полученные с листа</param>
        /// <param name="productLimitCommondBound">Уже установленные границы из БД Specs</param>
        private ProductLimit ProductLimitToWrtiteSpecs(ProductLimit prlimit, ProductLimitCommondBound productLimitCommondBound)
        {
            ProductLimit productLimit = new ProductLimit(prlimit.ProductionUnit_Desc, prlimit.Tag, prlimit.ProductCode, prlimit.Author);

            if (prlimit.HiHiRule) { productLimit.HiHi = prlimit.HiHi; }
            else { productLimit.HiHi = productLimitCommondBound.HiHi; }
            if  (prlimit.HiRule){productLimit.Hi = prlimit.Hi; }
            else { productLimit.Hi = productLimitCommondBound.Hi; }
            if  (prlimit.TargetRule) {productLimit.Target = prlimit.Target; }
            else { productLimit.Target = productLimitCommondBound.Target; }
            if  (prlimit.LoRule) {productLimit.Lo = prlimit.Lo; }
            else { productLimit.Lo = productLimitCommondBound.Lo; }
            if (prlimit.LoLoRule) { productLimit.LoLo = prlimit.LoLo; }
            else { productLimit.LoLo = productLimitCommondBound.LoLo; }
            productLimit.Author = prlimit.Author;

            return productLimit;
        }
             

        /// <summary>
        /// Обновление границы по одному тегу и одному продукту
        /// </summary>
        /// <param name="prlimit">Передаем класс ProductLimit сформированный из строки листа</param>
        /// <returns></returns>
        public int UpdateVarSpecsOneBound(ProductLimit prlimit, Int64 VS_ID, Int64 Product_Id)
        {
            int flagResult = 0;


            using (SqlConnection con = new SqlConnection(connString))
            {

                //Начинаю Транзанкцию
                //Не нужно на табл стоит тригер делающий тоже самое
                //SqlCommand cmdUpdate = new SqlCommand(String.Format("UPDATE VarSpecs SET Date_Expiration=GETDATE() WHERE VS_Id=@VS_Id"), con);
                SqlCommand cmdInsert = new SqlCommand(string.Format("INSERT INTO VarSpecs (Date_Effective, L_Reject, L_Warning, Target,U_Warning, U_Reject,Var_Id, Author, Product_Id  )" +
                    " VALUES (GETDATE(), @L_Reject, @L_Warning, @Target, @U_Warning, @U_Reject, @Var_Id, @Author, @Product_Id)"), con);


                SqlParameter paramTran = new SqlParameter();
                paramTran.ParameterName = "@L_Reject";
                if (prlimit.LoLo == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.LoLo; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@L_Warning";
                if (prlimit.Lo == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.Lo; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@Target";
                if (prlimit.Target == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.Target; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@U_Warning";
                if (prlimit.Hi == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.Hi; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@U_Reject";
                if (prlimit.HiHi == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.HiHi; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                //@Var_Id
                paramTran = new SqlParameter
                {
                    ParameterName = "@Var_Id",
                    Value = prlimit.Tag,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@Author
                paramTran = new SqlParameter
                {
                    ParameterName = "@Author",
                    Value = prlimit.Author,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@Product_Id
                paramTran = new SqlParameter
                {
                    ParameterName = "@Product_Id",
                    Value = Product_Id,
                    SqlDbType = System.Data.SqlDbType.BigInt
                };
                cmdInsert.Parameters.Add(paramTran);

                try
                {
                    con.Open();
                    if (cmdInsert.ExecuteNonQuery() == 1)
                    {
                        flagResult = 1;
                    }
                    else
                    {
                        flagResult = 0;
                    }
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                    // При возникновении любой ошибки выполняется откат транзакции.
                }
            }


            return flagResult;
        }

        /// <summary>
        /// Создание границы по одному тегу и одному продукты
        /// </summary>
        /// <param name="prlimit">Передаем класс ProductLimit сформированный из строки листа</param>
        /// <param name="entityProduct">Экземпляр класса EntityProduct (т.е. Продукта) для которого создаются новые границы</param>
        /// <returns></returns>
        public int CreateVarSpecsOneBound(ProductLimit prlimit, EntityProduct entityProduct)//
        {
            int flagResult = 0;

            using (SqlConnection con = new SqlConnection(connString))
            {


                //INSERT INTO [dbo].[VarSpecs]
                //([Date_Effective], [L_Reject], [L_Warning], [Target], [U_Warning], [U_Reject], [Author],[Var_Id],[Product_Id])
                //    select  GetDate(), 88,11,12,13,14,'valex','sinusoid1', [Product_Id] from[dbo].[Products] P
                //                                      left join[dbo].[ProductionUnits] pu ON pu.ProductionUnit_Id=p.ProductionUnit_Id
                //                                      where[Product_Code] = 'VVVVV' and[ProductionUnit_Desc]='OMC'

                SqlCommand cmdInsert = new SqlCommand(string.Format("INSERT INTO VarSpecs (Date_Effective, L_Reject, L_Warning, Target,U_Warning, U_Reject,Var_Id, Author, Product_Id  )" +
                " SELECT GETDATE(), @L_Reject, @L_Warning, @Target, @U_Warning, @U_Reject, @Var_Id, @Author, [Product_Id] from[dbo].[Products] P "+
                                                     " LEFT JOIN [dbo].[ProductionUnits] PU ON PU.ProductionUnit_Id=P.ProductionUnit_Id "+
                                                     " WHERE [Product_Code] = @Product_Code AND [ProductionUnit_Desc]=@ProductionUnit_Desc"), con);

                //                    " VALUES (GETDATE(), @L_Reject, @L_Warning, @Target, @U_Warning, @U_Reject, @Var_Id, @Author, @Product_Id)"), con);

                SqlParameter paramTran = new SqlParameter();
                paramTran.ParameterName = "@L_Reject";
                if (prlimit.LoLo == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.LoLo; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@L_Warning";
                if (prlimit.Lo == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.Lo; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);


                paramTran = new SqlParameter();
                paramTran.ParameterName = "@Target";
                if (prlimit.Target == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.Target; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@U_Warning";
                if (prlimit.Hi == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.Hi; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                paramTran = new SqlParameter();
                paramTran.ParameterName = "@U_Reject";
                if (prlimit.HiHi == null)
                { paramTran.Value = DBNull.Value; }
                else
                { paramTran.Value = prlimit.HiHi; }
                paramTran.SqlDbType = System.Data.SqlDbType.Float;
                cmdInsert.Parameters.Add(paramTran);

                //@Var_Id
                paramTran = new SqlParameter
                {
                    ParameterName = "@Var_Id",
                    Value = prlimit.Tag,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@Author
                paramTran = new SqlParameter
                {
                    ParameterName = "@Author",
                    Value = prlimit.Author,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@Product_Code
                paramTran = new SqlParameter
                {
                    ParameterName = "@Product_Code",
                    Value = entityProduct.Product_Code,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                //@ProductionUnit_Desc
                paramTran = new SqlParameter
                {
                    ParameterName = "@ProductionUnit_Desc",
                    Value = prlimit.ProductionUnit_Desc,
                    SqlDbType = System.Data.SqlDbType.NVarChar
                };
                cmdInsert.Parameters.Add(paramTran);

                try
                {
                    con.Open();


                    if (cmdInsert.ExecuteNonQuery() == 1)
                    {
                        flagResult = 1;
                    }
                    else
                    {
                        flagResult = 0;
                    }
                }
                catch (Exception e)
                {
                    lastError = e.Message;
                }



            } //Конец using с SQLConnection


            return flagResult;
        }




    }
}
