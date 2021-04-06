using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions; //Brings in Regex

namespace MeetingAddIn
{
    partial class FormRegion1
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("MeetingAddIn.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.

        private void populate_list() 
        {
            participant_list.Items.Add("Hank McCoy        $20/hr");
            participant_list.Items.Add("Scott Summers     $32/hr");
            participant_list.Items.Add("Charles Xavier    $18/hr");
            participant_list.Items.Add("Eric Lenshaw      $27/hr");
            participant_list.Items.Add("Jean Grey         $37/hr");
            participant_list.Items.Add("Johnny Cage       $40/hr");
            participant_list.Items.Add("Steve Rogers      $15/hr");
            participant_list.Items.Add("Kitty Pryde       $19/hr");
            participant_list.Items.Add("James Howlett     $21/hr");
            participant_list.Items.Add("Sean Cassidy      $33/hr");
            participant_list.Items.Add("Ororo Munroe      $38/hr");

        }
        double parsedNumber = 0;
        double total;

        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
            participant_list.Items.Clear();
            populate_list();
            

            //double parsedNumber = 0;
           // decimal total;
           

            if ((participant_list.SelectedItems.Count >= 1) && (!(string.IsNullOrWhiteSpace(duration_box.Text))) && (double.TryParse(duration_box.Text, out parsedNumber)))
            {
                for (int i = participant_list.Items.Count - 1; i >= 0; i--) 
                {
                    if (participant_list.GetSelected(i)) 
                    {
                        string temp = (string)participant_list.Items[i];

                        double result = 0;
                        double.TryParse(Regex.Match(temp, @"\d+").Value, out result);

                        total =+ result;
                        
                    }
                }

                calculate_button.Enabled = true;
            }
            else 
            {
                //calculate_button.Enabled = false;
            }
            
           
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
        }

        private void calculate_button_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            double meeting_cost = parsedNumber*total;
            output_box.Text = "$ " + 256;

        }
    }
}
