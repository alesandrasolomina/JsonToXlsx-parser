using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Newtonsoft.Json;

namespace JsonToXlsx_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            // getting json string 
            IEnumerable<string> lines = File.ReadLines("result.json");
            string json = $@"{String.Join("", lines)}";

            // using jarray from newtonsoft to get messages
            var dynamicObject = JsonConvert.DeserializeObject<dynamic>(json)!;
            var name = dynamicObject.name;
            var messagesJarray = dynamicObject.messages;

            //creating a list from JArray, for easier use of split for datetime
            var records = new List<Message>();
            foreach (var message in messagesJarray)
            {
                records.Add(new Message { From = message.from.ToString(), Date = message.date.ToString(), Text = message.text.ToString() });
            }

            // creating workbook
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            // setting headers
            IRow row1 = sheet1.CreateRow(0);
            row1.CreateCell(0).SetCellValue("From");
            row1.CreateCell(1).SetCellValue("Date");
            row1.CreateCell(2).SetCellValue("Time");
            row1.CreateCell(3).SetCellValue("Text");

            // wrting data from messages to workbook
            for (int i = 0; i < messagesJarray.Count; i++)
            {   
                // checking if 
                IRow row = sheet1.CreateRow(i + 1);

                string[] dt = records[i].Date.Split(' ');
                row.CreateCell(0).SetCellValue(records[i].From.ToString());
                row.CreateCell(1).SetCellValue(dt[0].ToString());
                row.CreateCell(2).SetCellValue(dt[1].ToString());
                if (records[i].Text.ToString() != "")
                {
                   row.CreateCell(3).SetCellValue(records[i].Text.ToString());
                }
                else
                {
                    row.CreateCell(3).SetCellValue("message contained no text");
                }

            }

                // saving workbook and closing
            FileStream sw = File.Create($"{name}.xlsx");
            workbook.Write(sw, false);
            sw.Close();
            Console.WriteLine("Json successfully parsed to xlsx");
            }
        }
    }

