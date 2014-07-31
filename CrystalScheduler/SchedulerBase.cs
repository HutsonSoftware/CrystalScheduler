using System.Windows.Forms;

namespace CrystalScheduler
{
    public partial class SchedulerBase : Form
    {
        internal string _title;

        public SchedulerBase()
        {
            InitializeComponent();
            _title = "Crystal Scheduler";
        }
        public string Title { get { return _title; } }
    }
}
