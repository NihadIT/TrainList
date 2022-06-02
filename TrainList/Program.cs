using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using OfficeOpenXml;
using TrainList.Models;

namespace TrainList
{
    internal class Program
    {
        static void Main(string[] args)
        {
            RootCollection roots = null;

            Console.Write("Номер состава: ");
            DocHandling.trainIndex = Console.ReadLine();

            //Читаем xml файл
            DocHandling.ReadXmlFile(ref roots);

            if (roots != null)
                //Создаем и записываем в excel файл
                DocHandling.CreateExcelFile();

            Console.Read();
        }
    }
}
