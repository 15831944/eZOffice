using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZx.AddinManager;
using eZx.RibbonHandler;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.Debug
{
    class EcTest4 : IExcexExCommand
    {
        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            try
            {
                DoSomething(excelApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }

        // 开始具体的调试操作
        private static void DoSomething(Application excelApp)
        {
            testSpeedMode(excelApp);
        }

        #region ---   具体的调试操作

        private static void testSpeedMode(Application excelApp)
        {
            DateTime[] Xdate = new DateTime[]
            {
                new DateTime(2013,10,17 ), new DateTime(2013,10,18 ), new DateTime(2013,10,19 ), new DateTime(2013,10,20 ), new DateTime(2013,10,21 ), new DateTime(2013,10,22 ), new DateTime(2013,10,23 ), new DateTime(2013,10,24 ), new DateTime(2013,10,25 ), new DateTime(2013,10,26 ), new DateTime(2013,10,27 ), new DateTime(2013,10,28 ), new DateTime(2013,10,29 ), new DateTime(2013,10,30 ), new DateTime(2013,10,31 ), new DateTime(2013,11,1 ), new DateTime(2013,11,2 ), new DateTime(2013,11,3 ), new DateTime(2013,11,4 ), new DateTime(2013,11,4 ), new DateTime(2013,11,5 ), new DateTime(2013,11,6 ), new DateTime(2013,11,7 ), new DateTime(2013,11,7 ), new DateTime(2013,11,8 ), new DateTime(2013,11,9 ), new DateTime(2013,11,10 ), new DateTime(2013,11,11 ), new DateTime(2013,11,12 ), new DateTime(2013,11,13 ), new DateTime(2013,11,14 ), new DateTime(2013,11,15 ), new DateTime(2013,11,16 ), new DateTime(2013,11,17 ), new DateTime(2013,11,18 ), new DateTime(2013,11,19 ), new DateTime(2013,11,20 ), new DateTime(2013,11,21 ), new DateTime(2013,11,22 ), new DateTime(2013,11,23 ), new DateTime(2013,11,24 ), new DateTime(2013,11,25 ), new DateTime(2013,11,26 ), new DateTime(2013,11,27 ), new DateTime(2013,11,28 ), new DateTime(2013,11,29 ), new DateTime(2013,11,30 ), new DateTime(2013,12,1 ), new DateTime(2013,12,2 ), new DateTime(2013,12,3 ), new DateTime(2013,12,4 ), new DateTime(2013,12,5 ), new DateTime(2013,12,5 ), new DateTime(2013,12,5 ), new DateTime(2013,12,6 ), new DateTime(2013,12,7 ), new DateTime(2013,12,8 ), new DateTime(2013,12,9 ), new DateTime(2013,12,10 ), new DateTime(2013,12,11 ), new DateTime(2013,12,12 ), new DateTime(2013,12,13 ), new DateTime(2013,12,14 ), new DateTime(2013,12,15 ), new DateTime(2013,12,16 ), new DateTime(2013,12,17 ), new DateTime(2013,12,18 ), new DateTime(2013,12,19 ), new DateTime(2013,12,20 ), new DateTime(2013,12,21 ), new DateTime(2013,12,22 ), new DateTime(2013,12,23 ), new DateTime(2013,12,24 ), new DateTime(2013,12,25 ), new DateTime(2013,12,26 ), new DateTime(2013,12,27 ), new DateTime(2013,12,28 ), new DateTime(2013,12,29 ), new DateTime(2013,12,30 ), new DateTime(2013,12,31 ), new DateTime(2014,1,1 ), new DateTime(2014,1,2 ), new DateTime(2014,1,3 ), new DateTime(2014,1,4 ), new DateTime(2014,1,5 ), new DateTime(2014,1,6 ), new DateTime(2014,1,7 ), new DateTime(2014,1,8 ), new DateTime(2014,1,9 ), new DateTime(2014,1,10 ), new DateTime(2014,1,11 ), new DateTime(2014,1,12 ), new DateTime(2014,1,13 ), new DateTime(2014,1,14 ), new DateTime(2014,1,15 ), new DateTime(2014,1,16 ), new DateTime(2014,1,23 ), new DateTime(2014,2,11 ), new DateTime(2014,2,20 ), new DateTime(2014,2,27 ), new DateTime(2014,3,6 ), new DateTime(2014,3,13 ), new DateTime(2014,3,15 ), new DateTime(2014,3,16 ), new DateTime(2014,3,17 ), new DateTime(2014,3,18 ), new DateTime(2014,3,18 ), new DateTime(2014,3,19 ), new DateTime(2014,3,19 ), new DateTime(2014,3,20 ), new DateTime(2014,3,21 ), new DateTime(2014,3,22 ), new DateTime(2014,3,23 ), new DateTime(2014,3,31 ), new DateTime(2014,4,7 ), new DateTime(2014,4,10 ), new DateTime(2014,4,11 ),
            };

            int[] XdateNum = new int[]
            {
                41564 ,41565 ,41566 ,41567 ,41568 ,41569 ,41570 ,41571 ,41572 ,41573 ,41574 ,41575 ,41576 ,41577 ,41578 ,41579 ,41580 ,41581 ,41582 ,41582 ,41583 ,41584 ,41585 ,41585 ,41586 ,41587 ,41588 ,41589 ,41590 ,41591 ,41592 ,41593 ,41594 ,41595 ,41596 ,41597 ,41598 ,41599 ,41600 ,41601 ,41602 ,41603 ,41604 ,41605 ,41606 ,41607 ,41608 ,41609 ,41610 ,41611 ,41612 ,41613 ,41613 ,41613 ,41614 ,41615 ,41616 ,41617 ,41618 ,41619 ,41620 ,41621 ,41622 ,41623 ,41624 ,41625 ,41626 ,41627 ,41628 ,41629 ,41630 ,41631 ,41632 ,41633 ,41634 ,41635 ,41636 ,41637 ,41638 ,41639 ,41640 ,41641 ,41642 ,41643 ,41644 ,41645 ,41646 ,41647 ,41648 ,41649 ,41650 ,41651 ,41652 ,41653 ,41654 ,41655 ,41662 ,41681 ,41690 ,41697 ,41704 ,41711 ,41713 ,41714 ,41715 ,41716 ,41716 ,41717 ,41717 ,41718 ,41719 ,41720 ,41721 ,41729 ,41736 ,41739 ,41740
            };
            double[] Y = new double[]
            {
                0     ,1.63  ,2.08  ,3.49  ,4.27  ,4.65  ,5.29  ,6.05  ,5.74  ,6.17  ,6.58  ,7.19  ,6.28  ,7.24  ,8.38  ,9.18  ,10.58 ,12.54 ,14.82 ,16.33 ,17.41 ,19.06 ,20.93 ,21.64 ,23.35 ,25.39 ,27.36 ,28.55 ,29.2  ,29.94 ,31.22 ,32.33 ,32.82 ,33.21 ,32.77 ,35.54 ,38.29 ,41.08 ,44.39 ,47.18 ,50    ,52.18 ,55.06 ,57.77 ,56.81 ,56.26 ,56.97 ,57.6  ,59.92 ,62.47 ,65.19 ,68.5  ,70.74 ,71.41 ,73.07 ,75.99 ,78.31 ,80.79 ,82.7  ,81.73 ,80.8  ,81.49 ,83.53 ,84.77 ,85.97 ,86.63 ,87.55 ,89.01 ,88.25 ,89.4  ,89.14 ,90.35 ,91.05 ,92.11 ,91.32 ,92.42 ,91.91 ,92.67 ,93.05 ,93.61 ,94.1  ,93.38 ,94.34 ,93.56 ,92.66 ,93.27 ,92.79 ,92.12 ,92.65 ,92.31 ,92.78 ,93.25 ,92.43 ,92.74 ,92.87 ,92.81 ,92.42 ,92.82 ,92.44 ,92.95 ,92.67 ,92.49 ,93.84 ,94.11 ,93.82 ,94.04 ,94.61 ,94.88 ,93.46 ,93.69 ,93.87 ,94.36 ,94.55 ,94.8  ,95.15 ,95.5  ,95.26 ,
            }; // 
            DebugUtils.show5();

            var ass = AppDomain.CurrentDomain.GetAssemblies();
            DebugUtils.ShowEnumerable(ass);


        }

        #endregion
    }
}