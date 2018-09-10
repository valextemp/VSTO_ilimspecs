-- Получение данных из настроечных файлов
        public void GetSetting()
        {
            string  sAttr = ConfigurationManager.AppSettings.Get("Key0");
            worksheet.Cells[1, 2] = sAttr;
            string cnStr=null;

            try
            {
                cnStr = ConfigurationManager.ConnectionStrings["specs1"].ConnectionString;
            }
            catch (Exception)
            {
                // без обработки
                
            }
            
            if (cnStr!=null) worksheet.Cells[2, 2] = cnStr;
            else worksheet.Cells[2, 2] = "Нет такой строки подключения";
        }
