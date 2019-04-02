using System;
using System.IO;
using Gma.QrCodeNet.Encoding;
using System.Drawing;
using System.Drawing.Imaging;
using Gma.QrCodeNet.Encoding.Windows.Render;


namespace QRcreate
{
    class QRImage
    {

        private StrJudge jud = new StrJudge();

       

        /// <summary>
        /// 文字转换成二维码格式并存储
        /// </summary>
        /// <param name="mark"></param>
        /// <param name="text"></param>
        private void saveQRImage(int mark, string text)
        {
            QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.M);
            QrCode qrCode = qrEncoder.Encode(text);

            //图片名称
            String imgname = mark.ToString("000") + text.Substring(0, 4) + ".png";
            GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);

            string str = System.Environment.CurrentDirectory;
            string Document = str + @"\picture";
            if (!Directory.Exists(Document)) //如果该文件夹不存在就建立这个新文件夹
            {
                Directory.CreateDirectory(Document);
            }

            using (MemoryStream ms = new MemoryStream())
            {
                renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms);
                Image img = Image.FromStream(ms);
                string path = Document + @"\" + imgname;// 文件的完全路径

                for(int i = 1; File.Exists(path); i++)
                {
                    string insert = "(" + i + ")";
                    string temp = path.Insert(path.Length - ".png".Length, insert);
                    if(!File.Exists(temp))
                    {
                        path = temp;
                        break;
                    }
                }

                img.Save(path);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(@"第 {0} 行字符串 {1} 生成二维码成功，存储路径为： {2}", mark, text, path);

            }
        }

        /// <summary>
        /// 将EXCEL文件单元格内字符串逐行转换为二维码
        /// </summary>
        /// <param name="excelPath"></param>
        public void excelToQR(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("EXCEL文件不存在: {0}", excelPath);
                return;
            }
            else
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(excelPath, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[1];

                
                string text = ((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[1, 1]).Text.ToString();

                for (int i = 1; text != string.Empty; i++)
                {
                    
                    if (jud.Judge(text) == 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("第{0}行单元格内字符生成失败：字符过长",i);
                    }
                    else if(jud.Judge(text) == 2)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("第{0}行单元格内字符生成失败：未知错误", i);
                        Console.WriteLine(text);
                    }
                    else
                    {
                        saveQRImage(i, text);
                    }
                    text = ((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[i + 1, 1]).Text.ToString();
                }

                app.Quit();
                PublicMethod.Kill(app);


                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("EXCEL文件生成完毕");
            }





}
        /// <summary>
        /// 将TEXT文件内容逐行转换为二维码
        /// </summary>
        /// <param name="textPath"></param>
        public void textToQR(string textPath)
        {
            if (!File.Exists(textPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("TEXT文件不存在: {0}", textPath);
                return;
            }
            else
            {
                string[] Contenets = File.ReadAllLines(textPath); //逐行读取文档内容
                int lines = Contenets.Length;



                for (int i = 0; i < lines; i++)
                {
                    if(jud.Judge(Contenets[i]) == 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("第{0}行字符串 {1} 生成失败：过长", i, Contenets[i]);
                    }
                    else if(jud.Judge(Contenets[i]) == 2)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("第{0}行字符串 {1} 生成失败：未知错误", i, Contenets[i]);
                    }
                    else
                    {
                        saveQRImage(i, Contenets[i]);
                    }
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(@"生成完毕");

            }
        }



        
        /// <summary>
        /// 在控制台用字符块表示二维码
        /// </summary>
        /// <param name="mark"></param>
        /// <param name="str"></param>
        public void blockQR(int mark, string str)
        {
            //字符串过长的情况
            if (jud.Judge(str) == 1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(@"第{0}个字符串 {1} 生成失败：字符串过长",mark,str);
                return;
            }

            //未知错误的情况
            else if(jud.Judge(str) == 2)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(@"第{0}个字符串 {1} 生成失败：出现未知错误",mark,str);
                return;
            }

            //正常输出
            else
            {
                QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.M);
                QrCode qrCode = qrEncoder.Encode(str);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(@"第{0}个字符串 {1} 生成成功：", mark, str);
                //字符块表示部分
                for (int i = 0; i < (qrCode.Matrix.Width + 2); i++) //将获得的二维码数组用黑白字块表示
                {
                    Console.BackgroundColor = ConsoleColor.White;//上边框
                    Console.Write("  ");
                }
                Console.WriteLine();
                for (int j = 0; j < qrCode.Matrix.Width; j++)
                {
                    Console.BackgroundColor = ConsoleColor.White;
                    Console.Write("  ");  //在左边添加边框
                    for (int i = 0; i < qrCode.Matrix.Height; i++)
                    {
                        if (qrCode.Matrix[i, j])
                        {
                            Console.BackgroundColor = ConsoleColor.Black;//黑色字块，实际显示的是空格字符的背景颜色
                            Console.Write("  ");
                        }
                        else
                        {
                            Console.BackgroundColor = ConsoleColor.White;//白色字块
                            Console.Write("  ");

                        }
                    }
                    Console.BackgroundColor = ConsoleColor.White;  //右边框
                    Console.Write("  ");
                    Console.WriteLine();
                }
                for (int i = 0; i < (qrCode.Matrix.Width + 2); i++)
                {
                    Console.BackgroundColor = ConsoleColor.White;//下边框
                    Console.Write("  ");
                }
                Console.WriteLine();
                Console.ResetColor();
                Console.WriteLine();

            }
        }
    }
}
