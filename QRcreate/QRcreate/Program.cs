using System;

namespace QRcreate
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("使用方法：");
            Console.WriteLine("1. 直接输入字符串生成二维码，字符串之间以空格分隔");
            Console.WriteLine("2. 输入‘-f’+ 'text文档地址'以逐行读取文档内容生成二维码， 多个文档地址以空格分隔");
            Console.WriteLine("3. 输入‘-e’+ 'excel文档地址'以逐行读取每行第一列单元格内容生成二维码， 多个文档地址以空格分隔");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("注意：生成二维码的字符串长度不能超过20");
            QRImage image = new QRImage();

            //获取命令行参数
            string[] argsArray = System.Environment.GetCommandLineArgs();

            int flag = 0, sign = 1;
            for(int i = 1; i < argsArray.Length; i++)
            {
                if(argsArray[i] == "-f")
                {
                    flag = 1;
                    sign = i + 1;
                    break;
                }
                else if(argsArray[i] == "-e")
                {
                    flag = 2;
                    sign = i + 1;
                    break;
                }
            }
            switch(flag)
            {
                case 0:
                    for(int i = sign; i < argsArray.Length; i++)
                    {
                        image.blockQR(i, argsArray[i]);
                    }
                    break;
                case 1:
                    for(int i = sign; i < argsArray.Length; i++)
                    {
                        image.textToQR(argsArray[i]);
                    }
                    break;
                case 2:
                    for (int i = sign; i < argsArray.Length; i++)
                    {
                        image.excelToQR(argsArray[i]);
                    }
                    break;
                default:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("未知问题");
                    break;
            }
            
            Console.ReadLine();

        }
    }
}
