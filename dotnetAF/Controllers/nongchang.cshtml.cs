using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace jingying.Pages
{
    public class nongchangModel : PageModel
    {
        public void OnGet()
        {

        }
        public static string tuiguang {
            get;

            set;
        }


        public static string licheng
        {
            get;

            set;
        }
        public static string tongshu
        {
            get;

            set;
        
        }
    }
}