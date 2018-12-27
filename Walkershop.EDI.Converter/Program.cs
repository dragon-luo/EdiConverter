using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using log4net;

namespace Walkershop.EDI.Converter
{
    static class Program
    {
        private static Mutex mutex = null;

        //日志记录器定义
        private static ILog logger = LogManager.GetLogger(typeof(Program));

        //常量定义
        private const string TIP_CONTENT = "请输入功能代码：";
        private const string HELP_CONTENT = "功能代码说明：\n1 - 采购收货数据导入 \n2 - 采购收货数据导出 \n3 - 出货数据导出 \n4 - 退出";
        private const string CONSOLE_TITLE = "EDI Data Converter";
        private const string ERROR_MESSAGE = "错误的功能代码 [{0}]";
                        
        static Program()
        {
            //设置控制台标题
            Console.Title = CONSOLE_TITLE;
        }


        [STAThread]      
        static void Main(string[] args)
        {
            bool firstInstance;
         
            //校验多用户环境下实例是否唯一
            mutex = new Mutex(true, @"Global\Walkershop.EDI.Converter", out firstInstance);
            try
            {
                if (!firstInstance)
                {
                    Console.WriteLine("已有实例运行，输入回车退出 .....");
                    Console.ReadLine();
                    return;
                }
                else
                {                    
                    RunConverter(args);
                }
            }
            catch (Exception ex) 
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
            }
            finally
            {
                //只有第一个实例获得控制权
                if (firstInstance)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Close();
                mutex = null;
            }
        }

        static void RunConverter(string[] args)
        {                
            string functionCode = string.Empty;

            if (args.Count() == 0)
            {
                args = new string[1];

                ConsoleWriteOutHelp();
                ConsoleWriteOutTip();
                functionCode = ConsoleReadFunctionCode();             
            }
            else
            {
                functionCode = args[0];
            }

            bool loop = true;

            //功能代码不正确，则输出提示信息
            while (loop)
            {
                loop = !IsFunctionCode(functionCode);

                if (loop)
                {
                    string errMessage = string.Format(ERROR_MESSAGE, functionCode);

                    logger.Error(errMessage);
                    Console.WriteLine(errMessage);
                    ConsoleWriteOutTip();
                    functionCode = ConsoleReadFunctionCode();
                }
            }

            //根据正确的功能代码运行相应功能
            EDIConverter edi = EDIConverter.GetInstance();
            edi.Run(int.Parse(functionCode));
        }
      
        static void ConsoleWriteOutHelp() 
        {
            Console.WriteLine(HELP_CONTENT);
        }

        static void ConsoleWriteOutTip()
        {
            Console.Write(TIP_CONTENT);
        }

        static void ConsoleWriteOutMessage(string errMessage)
        {
            Console.WriteLine(errMessage);
        }

        static string ConsoleReadFunctionCode() 
        {
            return Console.ReadLine().ToString().Trim();
        }

        static bool IsFunctionCode(string functionCode) 
        {               
            try
            {
                
                object value = int.Parse(functionCode);

                return Enum.IsDefined(typeof(FucnctionType), value);

            }catch
            {
                return false;
            }
        }      


    }
}
