using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace System.ComponentModel.DataAnnotations
{
    [AttributeUsage(AttributeTargets.Property)]
    public class PlaceHolder : Attribute
    {
        public string Text { get; set; }
        public PlaceHolder(string Text)
        {
            this.Text = Text;
        }
    }
}
