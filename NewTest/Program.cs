using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KompasAPI7;
using Kompas6API5;
using Kompas6Constants;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Pdf2d_LIBRARY;

namespace NewTest
{
    class Program
    {
        public static KompasObject CreateKompas()
        {
            KompasObject kompas = (KompasObject)CreateApplicationObject("KOMPAS.Application.5");
            if (kompas != null) return kompas;
            throw new SystemException("Проблема запуска Kompas, возможно приложение не установлено!");
        }

        //Получает экземпляр запущенного компаса
        public static KompasObject GetKompas()
        {
            KompasObject kompas = (KompasObject)GetApplicationObject("KOMPAS.Application.5");
            if (kompas != null) return kompas;
            throw new SystemException("Проблема подключения к Kompas!");
        }

        private static object CreateApplicationObject(string progId)
        {
            try
            {
                object obj = Activator.CreateInstance(Type.GetTypeFromProgID(progId) /*Type.GetTypeFromCLSID(new Guid("FBE002A6-1E06-4703-AEC5-9AD8A10FA1FA"))*/);
                
                return obj;
            }
            catch
            {
                return null;
            }
        }

        private static object GetApplicationObject(string progId)
        {
            try
            {
                object obj = null;
                try
                {
                    obj = Marshal.GetActiveObject(progId);
                    return obj;
                }
                catch
                {
                    obj = Activator.CreateInstance(Type.GetTypeFromProgID(progId)/*Type.GetTypeFromCLSID(new Guid("FBE002A6-1E06-4703-AEC5-9AD8A10FA1FA"))*/);
                    return obj;
                }
            }
            catch
            {
                return null;
            }
        }
        
        public static void getStamp(ksDocument2D IDocument2D, KompasObject _kompas, int columnNumber,string columnText) //Функция замены содержимого полей левой стороны штампа
        {
            ksStamp iStamp = (ksStamp)IDocument2D.GetStamp();
            if (iStamp != null)
            {
                //Console.WriteLine("Первый if");
                if (iStamp.ksOpenStamp() == 1)
                {
                    //Console.WriteLine("Второй if");
                    iStamp.ksColumnNumber(columnNumber);
                    ksTextLineParam iTextLineParam = (ksTextLineParam)_kompas.GetParamStruct(29);
                    if (iTextLineParam != null)
                    {
                        //Console.WriteLine("3 if");
                        iTextLineParam.Init();
                        iTextLineParam.style = 32768;
                        ksDynamicArray iTextItemArray = (ksDynamicArray)_kompas.GetDynamicArray(3);
                        if (iTextItemArray != null)
                        {
                            //Console.WriteLine("4 if");
                            ksTextItemParam iTextItemParam = (ksTextItemParam)_kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
                            if (iTextItemParam != null)
                            {
                                //Console.WriteLine("5 if");
                                iTextItemParam.Init();
                                iTextItemParam.iSNumb = 0;
                                iTextItemParam.s = columnText;
                                iTextItemParam.type = 0;
                                ksTextItemFont iTextItemFont = (ksTextItemFont)iTextItemParam.GetItemFont();
                                if (iTextItemFont != null)
                                {
                                    //Console.WriteLine("6 if");
                                    iTextItemFont.Init();
                                    iTextItemFont.SetBitVectorValue(4096, true);
                                    iTextItemFont.color = 0;
                                    iTextItemFont.height = 3.5;
                                    iTextItemFont.ksu = 1;
                                    iTextItemArray.ksAddArrayItem(-1, iTextItemParam);
                                    iTextLineParam.SetTextItemArr(iTextItemArray);
                                    iStamp.ksTextLine(iTextItemParam);
                                    //IDocument2D.ksTextLine(iTextItemParam);

                                }
                            }
                        }
                    }
                    iStamp.ksCloseStamp();
                }
            }
        }

        public static void getStampSecond(ksDocument2D IDocument2D, KompasObject _kompas, string columnText) // Функция замены содержимого поля названия организации в штампе
        {
            ksStamp iStamp = (ksStamp)IDocument2D.GetStamp();
            if (iStamp != null)
            {
                //Console.WriteLine("Первый if");
                if (iStamp.ksOpenStamp() == 1)
                {
                    //Console.WriteLine("Второй if");
                    iStamp.ksColumnNumber(9);
                    ksTextLineParam iTextLineParam = (ksTextLineParam)_kompas.GetParamStruct(29);
                    if (iTextLineParam != null)
                    {
                        //Console.WriteLine("3 if");
                        iTextLineParam.Init();
                        iTextLineParam.style = 32769;
                        ksDynamicArray iTextItemArray = (ksDynamicArray)_kompas.GetDynamicArray(3);
                        if (iTextItemArray != null)
                        {
                            //Console.WriteLine("4 if");
                            ksTextItemParam iTextItemParam = (ksTextItemParam)_kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
                            if (iTextItemParam != null)
                            {
                                //Console.WriteLine("5 if");
                                iTextItemParam.Init();
                                iTextItemParam.iSNumb = 0;
                                iTextItemParam.s = columnText;
                                iTextItemParam.type = 0;
                                ksTextItemFont iTextItemFont = (ksTextItemFont)iTextItemParam.GetItemFont();
                                if (iTextItemFont != null)
                                {
                                    //Console.WriteLine("6 if");
                                    iTextItemFont.Init();
                                    iTextItemFont.SetBitVectorValue(4096, true);
                                    iTextItemFont.color = 0;
                                    iTextItemFont.height = 7;
                                    iTextItemFont.ksu = 1;
                                    iTextItemArray.ksAddArrayItem(-1, iTextItemParam);
                                    iTextLineParam.SetTextItemArr(iTextItemArray);
                                    iStamp.ksTextLine(iTextItemParam);
                                    //IDocument2D.ksTextLine(iTextItemParam);

                                }
                            }
                        }
                    }
                    iStamp.ksCloseStamp();
                }
            }
        }

        [STAThread]// надо чтобы работало окно выбора директории
        static void Main(string[] args)
        {
            //Нормально работающий пример с запуском Компаса с exe
            KompasObject _kompas = CreateKompas();
            IApplication _kompasApi7 = (IApplication)_kompas.ksGetApplication7();
            _kompasApi7.Visible = false ;

            IDocuments Documents = _kompasApi7.Documents;


            //Окно выбора директории
            //OpenFileDialog ofd = new OpenFileDialog();
            //ofd.DefaultExt = ".cdw";

            //if (ofd.ShowDialog() == DialogResult.OK)
            //{
            //    string dir = Directory.GetCurrentDirectory();
            //    string [] curDir = Directory.GetFiles(dir);


            //    Console.WriteLine(curDir);
            //    Console.ReadLine();
            //    //System.Diagnostics.Process.Start(ofd.FileName);
            //}
           

            //Флаг конвертации
            bool convFlag=false;
            bool monoColorFlag = true;
            bool loopContinue = true;
            while (loopContinue)
            {
                Console.WriteLine("Конвертировать в PDF? \n Y - Да; \n N - Нет");

                //Проверка ответа пользователя
                string convPdf = Console.ReadLine();
                Console.WriteLine("\n");
                convPdf.ToLower();
                switch (convPdf)
                {
                    case "y":
                        convFlag = true;
                        Console.WriteLine("Цвет линий оставить ЧБ? \n Y - Да \n N - Нет ");
                        string color = Console.ReadLine();
                        Console.WriteLine("\n");
                        color.ToLower();
                        if (color == "y")
                        {
                            monoColorFlag = false;
                        }
                        loopContinue = false;
                        break;

                    case "n":
                        Console.WriteLine("Конвертация в PDF не проводится ");
                        loopContinue = false;
                        break;

                    default:
                        Console.Clear();
                        break;
                }
            }

            Console.WriteLine("Выберите папку с чертежами.");
            FolderBrowserDialog ofd = new FolderBrowserDialog();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine($"Выбрана папка: {ofd.SelectedPath}");
                Console.Write("Введите текст, который хотите отобразить в штампе (если нужны пустые поля нажмите ENTER): ");
                string columnText = Console.ReadLine();
                if (columnText=="")
                {
                       columnText =" ";
                }

                string[] curDir = Directory.GetFiles(ofd.SelectedPath,"*.cdw"); // получаем путь до папки
                foreach (string curFile in curDir)
                {
                    
                    Console.WriteLine($"Открываем: {curFile}");
                    IKompasDocument kompasDocuments = Documents.Open(curFile, false, false);
                    //IKompasDocument2D kompas_document_2d = (IKompasDocument2D)kompasDocuments;
                    ksDocument2D IDocument2D = _kompas.ActiveDocument2D();
                    
                    for (int i = 110; i <= 190; i++)
                    {
                        // Если надо избежать определенные поля штампа
                        //if((i==130)|| (i == 131))
                        //{
                        //    continue;
                        //}
                        //getStamp(IDocument2D, _kompas, i, Convert.ToString(i)); //для отслеживания номеров ячеек штампа

                        getStamp(IDocument2D, _kompas, i, columnText);
                    }
                    getStampSecond(IDocument2D, _kompas, columnText);
                    
                    //Если флаг активен, то конвертируем в pdf
                    if (convFlag)
                    {
                        Console.WriteLine("Конвертируем...");

                        IConverter Converter = _kompasApi7.Converter[_kompas.ksSystemPath(5) + "\\Pdf2d.dll"];

                        //Устанавливаем цвет линий
                        IPdf2dParam pdfParam = Converter.ConverterParameters(0);
                        if (monoColorFlag)
                        {
                            pdfParam.ColorType = 0;
                        }
                        
                        Converter.Convert(curFile, curFile.Remove(curFile.Length - 4) + ".pdf",0 , false);

                    }
                   
                    kompasDocuments.Save(); 
                    Console.WriteLine($"Сохраняем и закрываем: {curFile}");
                    Console.WriteLine("\n");
                    kompasDocuments.Close(0);
                }
                Console.WriteLine("Редактирование штампов чертежей прошло успешно!");
                Console.WriteLine("Убираем мусор...");

                //Удаляем bak файлы из папки
                string[] bakDir = Directory.GetFiles(ofd.SelectedPath, "*.bak");
                foreach (string bakFile in bakDir)
                {
                    FileInfo fileInf = new FileInfo(bakFile);
                    if (fileInf.Exists)
                    {
                        fileInf.Delete();
                    }
                }

                Console.WriteLine("Готово!");
                Console.Write("Нажмите Enter для завершения работы...");
                Console.ReadLine();
            }
            _kompas.Quit();
            _kompas = null;
            //_kompasApi7.Quit();
            //_kompasApi7 = null;
        }
    }
}
