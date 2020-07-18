using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AutoTest.Pages
{
    /// <summary>
    /// W001.xaml の相互作用ロジック
    /// </summary>
    public partial class W001 : Window
    {
        DispatcherTimer timer = new DispatcherTimer();
        private System.Drawing.Point _oldPoint;

        public W001()
        {
            InitializeComponent();

            timer.Tick += new EventHandler(timer1_Tick);
            timer.Start();
            _oldPoint = System.Windows.Forms.Cursor.Position;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            System.Drawing.Point point = System.Windows.Forms.Cursor.Position;

            if (point.X != _oldPoint.X || point.Y != _oldPoint.Y)
            {
                this.Close();
            }
        }
    }
}
