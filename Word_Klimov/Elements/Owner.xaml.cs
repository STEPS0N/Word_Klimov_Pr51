using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Word_Klimov.Elements
{
    /// <summary>
    /// Логика взаимодействия для Owner.xaml
    /// </summary>
    public partial class Owner : UserControl
    {
        public Owner(Context.OwnerContext roomOwner)
        {
            InitializeComponent();
            NameOwner.Content = $"{roomOwner.LastName} {roomOwner.FirstName} {roomOwner.SurName}";
        }
    }
}
