using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace QRcreate
{
    class StrJudge
    {
        /// <summary>
        /// 判断字符串是否符合自定义规则
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public int Judge(string str)
        {
            //字符串规则定义板块单独分出，便于修改
            //此处定义的规则是字符串长度在20以内
            if(Regex.IsMatch(str, @"^.{0,20}$")) //使用了正则表达式
            {
                return 0;
            }
            else if(str.Length > 20)
            {
                return 1;
            }
            else
            {
                return 2;
            }            
        }

    }
}
