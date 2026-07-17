using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace AutoBCC
{
    public class AddInModel : AModel<AddInModel>
    {
        public AddInModel() : base() { }
        public string Mail {
            get => GetValue<string>();
            set => SetValue(value);
        }
        public bool Enabled
        {
            get => GetValue<bool>();
            set => SetValue(value);
        }
    }
}
