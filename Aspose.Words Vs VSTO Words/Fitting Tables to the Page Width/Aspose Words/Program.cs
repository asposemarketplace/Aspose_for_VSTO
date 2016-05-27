﻿using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "";
            Document doc = new Document();
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                table.PreferredWidth = PreferredWidth.FromPercent(100);
            }
        }
    }
}
