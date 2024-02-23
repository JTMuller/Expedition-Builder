using Expedition_Builder_Online.Properties;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Schema;
using SColor = System.Drawing.Color; // Because Google API also uses Color
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System.Threading;
using static Expedition_Builder_Online.TestScreen;
//using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace Expedition_Builder_Online
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        public class GSheets
        {
            public static string Code = "1XcZ5V1G2dUMcm5qjYN1QHUaH1m6CaAA4rYKwDi6rfcE"; // Copied from Google Sheet
            public static string TabChar = "CharacterData"; // Tabname as named in Google Sheet
            public static string TabCharRange = "!B:CP"; // Columnrange

            public static string TabItem = "ItemData";
            public static string TabItemRange = "!A:AI";

            public static string TabStats = "CharacterStats";
            public static string TabStatsRange = "!B:CM";

            public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
            public static SheetsService Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    new ClientSecrets // obtained on Google Cloud Console API
                    {
                        ClientId = "876868570518-5pfb5gcispkc43f2n2c1ehlj7so691qk.apps.googleusercontent.com",
                        ClientSecret = "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"
                    },
                Scopes,
                "user",
                CancellationToken.None,
                new FileDataStore("MyAppsToken")
                ).Result,
                ApplicationName = "Google Sheets .NET API Quickstart",
            });
        }

        //
        //          Initial Load of Values
        //

        private void Initial_Load(object sender, EventArgs e)
        {
            TT_Gear_Load();
            TT_Talent_Load();

            CbCharSource.Items.Clear();
            for(int i = 0; i < Source.SubName.Count; i++)
            {
                CbCharSource.Items.Add(Source.Name(i, false));
            }

            Character_Data[0] = 100; // Role Rank
            Character_Data[1] = 100; // Skill Rank
            Character_Data[2] = 100; // Quest Rank
        }

        int Tab_Current = 1; // The current Tab in the form
        bool GearSelectiveSearch = false;
        int GearSelectiveSlot = 0;

        int[] AffinityChoice = new int[6] { 4, 4, 3, 3, 3, 3 }; // Saving the Affinity choices
        int[] AbilityChoice = new int[9]; // Saving the Ability choices
        int[] TalentChoice = new int[60]; // Saving the Talent choices

        int AbilityGrowthPoints = 30; // 30 start
        int AbilityGrowthPointsMax = 30; // 30 standard

        int TalentPoints = 60;
        int TalentPointsMax = 60;

        double[] AffinityStats = new double[19]; // excludes abilities
        double[] AbilityStats = new double[9];
        double[] GearStats = new double[28]; 
        double[] TalentStats = new double[30]; // includes two ranks

        ToolTip[] TT_Character = new ToolTip[31]; // 28 Tooltips, and 3 extra for source, save and load
        ToolTip[] TT_Affinity = new ToolTip[6]; // 6 Tooltips, not exclusive for each button
        ToolTip[] TT_Gear = new ToolTip[12]; // 12 Tooltips, for each gearpiece
        bool TT_Char_Check = false;
        bool TT_Affinity_Check = false;
        bool HelpMode = false;

        string Character_Name = "No Character Loaded";
        string Character_Code = "Enter Character Code";
        int Character_Data_Row = 0;
        int[] Character_Data = new int[91];
        double[] Character_Stats = new double[28];
        string[] Character_Stats_Print = new string[28];
        void Character_Load(string InputCode)
        {
            string ConvertCode = InputCode.ToLower();

            int Skipper = 0;

            try
            {
                var GValues = GSheets.Service.Spreadsheets.Values.Get(GSheets.Code, GSheets.TabChar + GSheets.TabCharRange).Execute().Values;

                foreach (var GRow in GValues)
                {
                    Skipper++;

                    if (GRow[0] != null && Skipper > 2)
                    {
                        if (GRow[0].ToString() == ConvertCode) // If the code matches
                        {
                            Character_Code = GRow[0].ToString();
                            Character_Name = GRow[1].ToString();

                            for (int i = 0; i <= 90; i++)
                            {
                                try
                                {
                                    string Row = GRow[i + 2].ToString();
                                    Character_Data[i] = int.Parse(Row);
                                }
                                catch
                                {
                                    Character_Data[i] = 0;
                                }
                            }
                            Character_Data_Row = Skipper;
                            break;
                        }
                    }

                }

            }
            catch
            {
                //MessageBox.Show(ex.Message, "Character Load did not work at Row " + Skipper.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("That input code did not work, maybe you typed it wrong (with a spacebar added where it shouldn't) or your connection with Sheets have been lost. A button has been added to refresh the Google Sheets connection!", "Character Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        void Character_Save(string Tab, string Range, bool Fresh = false)
        {
            string Placement = "";
            if (Fresh)
            {
                SpreadsheetsResource.ValuesResource.GetRequest GRequest = GSheets.Service.Spreadsheets.Values.Get(GSheets.Code, Tab + Range); // Tab + TabChar
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender2, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) { return true; };
                ValueRange GResponse = GRequest.Execute();
                IList<IList<Object>> GValues = GResponse.Values;
                Placement = (GValues.Count + 1).ToString();
            }
            else
            {
                Placement = Character_Data_Row.ToString();
            }
            string RangeFrom = Range.Substring(0, 2);
            string RangeTo = Range.Substring(2, 3);
            var GRange = $"{Tab}" + RangeFrom + Placement + RangeTo + Placement;
            var GValueRange = new ValueRange();

            GValueRange.Values = new List<IList<Object>>();
            List<Object> GValueRangeInner = new List<Object>();
            if ((Fresh && Tab == GSheets.TabChar) || (Tab == GSheets.TabStats)) // first editions need to be placed for fresh characters
            {
                GValueRangeInner.Add(Character_Code);
                GValueRangeInner.Add(Character_Name);
            }
            if (Fresh && Tab == GSheets.TabChar)
            {
                for (int i = 0; i <= 90; i++)
                {
                    GValueRangeInner.Add(Character_Data[i]);
                }
            }
            else if (Tab == GSheets.TabChar)
            {
                for (int i = 3; i <= 90; i++)
                {
                    GValueRangeInner.Add(Character_Data[i]);
                }
                RangeFrom = "!G";
                GRange = $"{Tab}" + RangeFrom + Placement + RangeTo + Placement;
            }
            else if (Tab == GSheets.TabStats)
            {
                for (int i = 0; i <= 87; i++) // other range, talents still added
                {
                    if (i <= 27)
                    {
                        GValueRangeInner.Add(Character_Stats_Print[i]); // Stats for Raw stats, StatsPrint for converted stats
                    }
                    else
                    {
                        GValueRangeInner.Add(Character_Data[i + 3]); // make up for the difference
                    }
                }
            }

            GValueRange.Values.Add(GValueRangeInner);

            var GUpdateRequest = G.Service.Spreadsheets.Values.Update(GValueRange, GSheets.Code, GRange);
            GUpdateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var GUpdateResponse = GUpdateRequest.Execute();
        }

        ///
        //          Working the Tabs
        //

        private void Tab_Click(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int Nr = int.Parse(TempBox.Name.Substring(5, 1));

            if (Tab_Current != Nr)
            {
                Form.ActiveForm.BackColor = Screen_Colors.ElementAt(Nr - 1); // To make the transition look slightly more okay

                for (int i = 1; i <= 5; i++)
                {
                    string PanelName = "Pnl" + i.ToString();
                    string PbTabName = "PbTab" + i.ToString();

                    PictureBox ThisBox = this.Controls.Find(PbTabName, true).FirstOrDefault() as PictureBox;
                    Panel ThisPanel = this.Controls.Find(PanelName, true).FirstOrDefault() as Panel;

                    if (i == Nr)
                    {
                        ImageChanger(ThisBox, "Tab-" + i.ToString() + "-ActiveS");
                        ThisPanel.Visible = true;
                        ImageChanger(PbBG, "Screen-" + Nr.ToString() + "S");
                    }
                    else if (i == Tab_Current)
                    {
                        ImageChanger(ThisBox, "Tab-" + i.ToString() + "-DeactiveS");
                        ThisPanel.Visible = false;
                    }
                }
                Tab_Current = Nr;
            }

            if (Tab_Current == 1)
            {
                Affinity_Calculate();
                Ability_Calculate(); // Calculates abilities and prints them
                Talent_Calculate();
                Talent_Calculate_TradeOff();
                SourceLoad(Character_Data[3]);
                Character_Data_Print();
                TT_Character_Load();
            }
            if (Tab_Current == 4) // Gear is activated
            {
                TempGearListReset();
            }
        }
        private void Tab_Hover(object sender, MouseEventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int Nr = int.Parse(TempBox.Name.Substring(5, 1));

            for (int i = 1; i <= 5; i++)
            {
                string PbTabName = "PbTab" + i.ToString();
                PictureBox ThisBox = this.Controls.Find(PbTabName, true).FirstOrDefault() as PictureBox;

                if (i != Tab_Current)
                {
                    if (i == Nr)
                    {
                        ImageChanger(ThisBox, "Tab-" + i.ToString() + "-SelectS");
                    }
                    else
                    {
                        ImageChanger(ThisBox, "Tab-" + i.ToString() + "-DeactiveS");
                    }
                }
            }
        }

        List<SColor> Screen_Colors = new List<SColor>()
        {
            SColor.FromArgb(0,126,255),
            SColor.FromArgb(109,161,124),
            SColor.FromArgb(94,255,117),
            SColor.FromArgb(214,59,194),
            SColor.FromArgb(46,46,46)
        };

        void ImageChanger(PictureBox PBox, string ResourceName)
        {
            if (PBox.Image != null)
            { PBox.Image.Dispose(); }
            PBox.Image = (Bitmap)Expedition_Builder_Online.Properties.Resources.ResourceManager.GetObject(ResourceName);
        }      

        //
        //          Dragging the Forms
        // 

        public static class Drag
        {
            public const int BtnDwn = 0xA1;
            public const int Cap = 0x2;

            [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
            public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
            [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
            public static extern bool ReleaseCapture();
        }
        private void Drag_Form(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Drag.ReleaseCapture();
                Drag.SendMessage(Handle, Drag.BtnDwn, Drag.Cap, 0);
            }
        }

        //
        //          Who Are You? (Tab 1)
        //

        private void Char_Save(object sender, EventArgs e)
        {
            Character_Data_Print();
            Character_Data_Return();
            if (Character_Data_Row != 0 && TxtCode.Text.ToLower() == Character_Code)
            {
                MessageBox.Show("Time to Save this character");
                Character_Save(GSheets.TabChar, GSheets.TabCharRange);
                Character_Save(GSheets.TabStats, GSheets.TabStatsRange);
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show(string.Format("{1} does not exist yet in the database.{0}Do you wish to create it with the following code?{0}Code: {2}", Environment.NewLine, Character_Name, TxtCode.Text.ToLower()), "Fresh Character", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    Character_Code = TxtCode.Text.ToLower();
                    MessageBox.Show("Great! Remember to ask the Expedition Master to link a ranking track when you start using this character!","You got it");
                    Character_Save(GSheets.TabChar, GSheets.TabCharRange, true);
                    Character_Save(GSheets.TabStats, GSheets.TabStatsRange, true);

                    TT_Character_Load();
                }
                if (dialogResult == DialogResult.No)
                {
                    
                }
            }
        }

        private void Char_Load(object sender, EventArgs e)
        {
            Character_Load(TxtCode.Text); // Load the data in
            Character_Data_Convert(); // Use the data and put it in the right buckets

            Affinity_Calculate();
            Gear_Calculate();
            Ability_Calculate(); // Calculates abilities and prints them
            Talent_Calculate();
            Talent_Calculate_TradeOff();

            LblCharName.Text = Character_Name;
            LblChar0.Text = Character_Data[0].ToString();
            LblChar1.Text = Character_Data[1].ToString();
            LblChar2.Text = Character_Data[2].ToString();
            SourceLoad(Character_Data[3]);
            Character_Data_Print();
            TT_Affinity_Load();
            TT_Character_Load();
        }

        void Char_Update_Silent(string Code)
        {
            Character_Load(Code);
            Character_Data_Convert();

            Affinity_Calculate();
            Gear_Calculate();
            Ability_Calculate(); // Calculates abilities and prints them
            Talent_Calculate();
            Talent_Calculate_TradeOff();

            Character_Data_Print();
            Character_Data_Return();
            
            Character_Save(GSheets.TabStats, GSheets.TabStatsRange);
        }

        void Char_Update_MainCast()
        {
            Char_Update_Silent("antonio");
            Char_Update_Silent("andrew");
            Char_Update_Silent("theresa");
            Char_Update_Silent("chimbie");
            Char_Update_Silent("leira");
            Char_Update_Silent("john");
            Char_Update_Silent("erebus");
            Char_Update_Silent("rilies");
            Char_Update_Silent("pearl");
        }

        /*
 0: Role
         * 0:Skill
 * 1:Quest
 * 2:Source
 * 3:Health Points
 * 4:Resource Points
 * 5:Link Points
 * 6:Movement Points
 * 7:Physical Prowess
 * 8:Physical Power
 * 9:Armor
 * 10:Avoidance
 * 11:Magical Prowess
 * 12:Magical Power
 * 13:Warding
 * 14:Resistance
 * 15:Healing Prowess
 * 16:Healing Power
 * 17:Ease
 * 18:Attune
 * 19:Precison
 * 20:Critical Success
 * 21:Critical Failure
 * 22:Athletics
 * 23:Control
 * 24:Perception
 * 25:Knowledge
 * 26:Magica
 * 27:Survival
 * 28:Deception
 * 29:Diplomacy
 * 30:Mentality
 */

        /*
         * Chardata
         * 3 base elements
         * 6 affinities
         * 9 abilities
         * 12 items
         * 60 talents
         */
        void Character_Data_Return()
        {
            for (int i = 4; i <= 9; i++)
            {
                /*
                 * 0-2 rank
3 source
4-9 affinity
10-18 ability
19-30 gear
30-89 talent
                 */
                Character_Data[i] = AffinityChoice[i - 4];
            }
            for (int i = 10; i <= 18; i++)
            {
                Character_Data[i] = AbilityChoice[i - 10];
            }
            for (int i = 19; i <= 30; i++)
            {
                Character_Data[i] = Gear.Stats[i - 19][0];
            }
            for (int i = 31; i <= 90; i++)
            {
                Character_Data[i] = TalentChoice[i - 31];
            }
        }

        void Character_Data_Convert()
        {

            // Load Affinity Scores
            for (int i = 1; i <= 6; i++)
            {
                Affinity_Change(i, Character_Data[i + 3]); // 4-9 are affinities
            }

            // Load Ability Growth Points
            AbilityGrowthPointsMax = 30; // base 30
            for (int i = 1; i <= 9; i++)
            {
                AbilityChoice[i - 1] = -1; // 10-18 are abilities
                Ability_Change(i, Character_Data[i + 9]);
            }

            // Load Equiped Gear
            TheNakedTruth(); // First, remove all gear
            for (int i = 1; i <= 12; i++)
            {
                if (Character_Data[i + 18] > 0)
                {
                    GearEquip(i, Character_Data[i + 18]); // 19-30 are gear itemcodes
                }
            }
            TT_Gear_Load();

            // Load Talents
            Talent_Reset();
            Talent_Point_Calculate();
            for (int i = 1; i <= 6; i++)
            {
                int Select = 0;
                if (i <= 5)
                {
                    for (int j = 0; j <= 10; j++)
                    {
                        Select = Talent_Select_Index(i, j);
                        if (Character_Data[Select + 31] > 0) // 31-90 are Talents
                        {
                            Talent_Selected(i, j, Select, true);
                        }
                        else
                        {
                            Talent_Selected(i, j, Select, false, false);
                        }
                    }
                }
                else
                {
                    Select = Talent_Select_Index(i, 0);
                    if (Character_Data[Select + 31] > 0)
                    {
                        for (int j = 0; j < Character_Data[Select + 31]; j++)
                        {
                            Talent_Selected_Ability(true);
                        }
                    }
                    else
                    {
                        Talent_Selected_Ability(false, false);
                    }

                }
            }
            Ability_Growth_Update();
            Ability_Visibility();
            Talent_Label_Update();
        }
        void Character_Data_Print()
        {
            for (int i = 0; i <= 27; i++)
            {
                if (i <= 18)
                {
                    Character_Stats[i] = AffinityStats[i] + GearStats[i] + TalentStats[i];
                }
                else
                {
                    Character_Stats[i] = AbilityStats[i-19] + GearStats[i] + TalentStats[i];
                }
                string LblName = "LblChar" + (i +4 ).ToString();
                Label ThisLbl = this.Controls.Find(LblName, true).FirstOrDefault() as Label;

                switch (i)
                {
                    case 15:
                        Character_Stats_Print[i] = (12 - Character_Stats[i]).ToString();
                        break;
                    case 17:
                        Character_Stats_Print[i] = (20 - Character_Stats[i]).ToString();
                        break;
                    case 3: case 18:
                        Character_Stats_Print[i] = (1 + Character_Stats[i]).ToString();
                        break;
                    default:
                        Character_Stats_Print[i] = Character_Stats[i].ToString();
                        break;
                }
                ThisLbl.Text = Character_Stats_Print[i];
            }        
        }
        private void TxtCode_TextChanged(object sender, EventArgs e)
        {
            if(TxtCode.Text == "UpdateAll")
            {
                Char_Update_MainCast();
            }

            if (TT_Char_Check)
            { TT_Character_Load(); }         
        }
        void TT_Character_Load()
        {
            string PbCharName = null;
            PictureBox PbC = null;
            string[] Extra = new string[2] { "Save", "Load" };

            for (int i = 0; i <= 27; i++)
            {
                if (TT_Char_Check)
                { TT_Character[i].Dispose(); }
                TT_Character[i] = new ToolTip();

                TT_Character[i].ToolTipTitle = Gear.StatNames[i+3];
                string TT = string.Format("Your {1} values originate from:{0}", Environment.NewLine, Gear.StatNames[i + 3]);

                if (i <= 18)
                {
                    // Aff, Gear, Tal
                    TT += string.Format("Affinity Growth: {1}.{0}", Environment.NewLine, AffinityStats[i]);
                }
                else
                {
                    // Ab i-19, Gear, Tal
                    TT += string.Format("Ability Growth: {1}.{0}", Environment.NewLine, AbilityStats[i-19]);

                }
                TT += string.Format("Gear: {1}.{0}", Environment.NewLine, GearStats[i]);
                TT += string.Format("Talents: {1}.{0}", Environment.NewLine, TalentStats[i]);


                PbCharName = "PbChar" + (i + 4).ToString();
                PbC = this.Controls.Find(PbCharName, true).FirstOrDefault() as PictureBox;
                TT_Character[i].SetToolTip(PbC, TT);
            }
            for (int i = 0; i <= 1; i++)
            {
                int y = i + 28;

                if (TT_Char_Check)
                { TT_Character[y].Dispose(); }
                TT_Character[y] = new ToolTip();

                TT_Character[y].ToolTipTitle = Extra[i] + " a Character";
                string TT = string.Format("Make sure you use the correct!{0}Current Code Sender: {1}", Environment.NewLine, TxtCode.Text);
                if (i == 0)
                { 
                    TT += string.Format("{0}Current Code in Data: {1}", Environment.NewLine, Character_Code); 
                    if (TxtCode.Text.ToLower() == Character_Code)
                    {
                        TT += string.Format("{0}Your current codes match!{0}Data will be written to the same character.", Environment.NewLine);
                    }
                    else
                    {
                        TT += string.Format("{0}Your current codes do not match!{0}Data will be written to a new character with the code {1}.", Environment.NewLine, TxtCode.Text.ToLower());
                    }
                
                }
                PbCharName = "Pb" + Extra[i];
                PbC = this.Controls.Find(PbCharName, true).FirstOrDefault() as PictureBox;
                TT_Character[y].SetToolTip(PbC, TT);
            }
            TT_Char_Check = true;
        }

        //
        //          Affinities (Tab 2)
        //

        private void Affinity_Activate(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int NrAff = int.Parse(TempBox.Name.Substring(5, 1)); // Affinity Place
            int NrAffS = int.Parse(TempBox.Name.Substring(7, 1)); // Which Spot

            Affinity_Change(NrAff, NrAffS);
            Affinity_Calculate();

            TT_Affinity_Load();
        }
        void Affinity_Change(int NrAffinity, int NrAffinitySelect)
        {
            int Skip = 0;
            if (NrAffinity == 2) { Skip = 7; }
            else if (NrAffinity > 2) { Skip = 14; }

            if (AffinityChoice[NrAffinity - 1] != NrAffinitySelect)
            {
                int Total = 5; // Duo affinity
                if (NrAffinity <= 2)
                { Total = 7; }

                for (int i = 1; i <= Total; i++) // There are 7 nodes for tri-affinities, and 5 nodes for duo-affinities.
                {
                    string PbAffName = "PbAff" + NrAffinity.ToString() + "_" + i.ToString(); // PbAff1_4 for example
                    PictureBox ThisBox = this.Controls.Find(PbAffName, true).FirstOrDefault() as PictureBox;

                    if (i == NrAffinitySelect)
                    {
                        ImageChanger(ThisBox, Icon_Affinity(i + Skip, true, NrAffinity));
                    }
                    else if (i == AffinityChoice[NrAffinity - 1])
                    {
                        ImageChanger(ThisBox, Icon_Affinity(i + Skip, false, NrAffinity));

                    }
                }
                AffinityChoice[NrAffinity - 1] = NrAffinitySelect;
            }
        }

        void Affinity_Calculate()
        {
            /* OBSOLUTE DUE TO START VALUES
            for (int i = 1; i <= 6; i++)
            {
                if (AffinityChoice[i - 1] == 0)
                {
                    if (i < 3)
                    {
                        Affinity_Change(i, 4);
                    }
                    else
                    {
                        Affinity_Change(i, 3);
                    }

                }
            } // Always choose a middle feat when calculating with empty slots
            */
            double SkillRank = (double)Character_Data[1] + TalentStats[29];

            // Offense Affinity
            double[] Off = Affinity_Results.ElementAt(AffinityChoice[0] + AIndexBoost[0]); // 1 Tri
            AffinityStats[4] = Math.Floor(SkillRank * Off[0]); // Physical Prowess
            AffinityStats[8] = Math.Floor(SkillRank * Off[1]); // Magical Prowess
            AffinityStats[12] = Math.Floor(SkillRank * Off[2]); // Healing Prowess

            // Defense and Survival Affinity
            double[] Def = Affinity_Results.ElementAt(AffinityChoice[1] + AIndexBoost[1]); // 2 Tri
            double[] DefF = Affinity_TalentBoost();
            double[] Sur = Affinity_Results.ElementAt(AffinityChoice[3] + AIndexBoost[3]); // 4
            double[] SurF = Affinity_TalentMod(1, 3);

            AffinityStats[6] = Math.Floor(SkillRank * Def[0] * DefF[0] * Sur[0] * SurF[0]); // Armor
            AffinityStats[7] = Math.Floor(SkillRank * Def[0] * DefF[0] * Sur[1] * SurF[1]); // Avoidance

            AffinityStats[10] = Math.Floor(SkillRank * Def[1] * DefF[1] * Sur[0] * SurF[0]); // Warding
            AffinityStats[11] = Math.Floor(SkillRank * Def[1] * DefF[1] * Sur[1] * SurF[1]); // Resistance

            AffinityStats[14] = Math.Floor(SkillRank * Def[2] * DefF[2] * Sur[0] * SurF[0]); // Ease
            AffinityStats[15] = Math.Floor(SkillRank * Def[2] * DefF[2] * Sur[1] * SurF[1]); // Attune

            // Strike
            double[] Str = Affinity_Results.ElementAt(AffinityChoice[2] + AIndexBoost[2]); // 3
            double[] StrF = Affinity_TalentMod(0, 2);
            AffinityStats[5] = Math.Floor(SkillRank * Str[0] * StrF[0]); // Physical Power
            AffinityStats[9] = Math.Floor(SkillRank * Str[0] * StrF[0]); // Magical Power
            AffinityStats[13] = Math.Floor(SkillRank * Str[0] * StrF[0]); // Healing Power
            AffinityStats[16] = Math.Floor(SkillRank * Str[1] * StrF[1]); // Precision

            // Endurance
            double[] End = Affinity_Results.ElementAt(AffinityChoice[4] + AIndexBoost[4]); // 5
            double[] EndF = Affinity_TalentMod(2, 4);
            AffinityStats[0] = Math.Floor(SkillRank * End[0] * EndF[0]); // Health Points
            AffinityStats[1] = Math.Floor(SkillRank * End[1] * EndF[1]); // Resource Points

            // Flexibility
            double[] Flx = Affinity_Results.ElementAt(AffinityChoice[5] + AIndexBoost[5]); // 6
            double[] FlxF = Affinity_TalentMod(3, 5);
            AffinityStats[2] = Math.Floor(SkillRank * Flx[0] * FlxF[0]); // Link Points
            AffinityStats[3] = Math.Floor(SkillRank * Flx[1] * FlxF[1]); // Movement Points
        }

        double[] Affinity_TalentMod(int T, int A)
        {
            double[] Aff = new double[2] { 1, 1 };

            if (TalentChoice[T] > 0 && TalentChoice[T + 4] > 0)
            {
                Aff[0] = 1.2;
                Aff[1] = 1.2;
            }
            else if (TalentChoice[T] > 0)
            {
                if (AffinityChoice[A] > 3)
                {
                    Aff[1] = 1.2;
                }
                else if (AffinityChoice[A] < 3)
                {
                    Aff[0] = 1.2;
                }
                else
                {
                    Aff[0] = 1.1;
                    Aff[1] = 1.1;
                }
            }
            else if (TalentChoice[T + 4] > 0)
            {
                if (AffinityChoice[A] > 3)
                {
                    Aff[0] = 1.2;
                }
                else if (AffinityChoice[A] < 3)
                {
                    Aff[1] = 1.2;
                }
                else
                {
                    Aff[0] = 1.1;
                    Aff[1] = 1.1;
                }
            }

            return Aff;
        }
        double[] Affinity_TalentBoost()
        {
            double[] Aff = new double[3] { 1, 1, 1 };

            if (TalentChoice[8] > 0)
            { Aff[0] = 1.2; }
            if (TalentChoice[9] > 0)
            { Aff[1] = 1.2; }
            if (TalentChoice[10] > 0)
            { Aff[2] = 1.2; }

            return Aff;
        }

        List<double[]> Affinity_Results = new List<double[]>()
        {
            new double[] { 0.25, 0.05, 0.05 },      // Top Left     (Max Red)
            new double[] { 0.15, 0.15, 0.05 },      // Top Middle   (Red / Blue)
            new double[] { 0.05, 0.25, 0.05 },      // Top Right    (Max Blue)
            new double[] { 0.12, 0.12, 0.12 },      // Middle       (Neutral)
            new double[] { 0.15, 0.05, 0.15 },      // Lower Left   (Red / Yellow)
            new double[] { 0.05, 0.15, 0.15 },      // Lower Right  (Blue / Yellow)
            new double[] { 0.05, 0.05, 0.25 },      // Lower        (Max Yellow)

            new double[] { 1.00, 0.00 },
            new double[] { 0.75, 0.12 },
            new double[] { 0.50, 0.25 }, // Armor (0,5) vs Avoid (1)
            new double[] { 0.25, 0.37 },
            new double[] { 0.00, 0.50 },

            new double[] { 0.20, 0.00 },
            new double[] { 0.15, 0.05 },
            new double[] { 0.10, 0.10 }, // Power (0,5) vs Precision (1), but evened due to more power
            new double[] { 0.05, 0.15 }, 
            new double[] { 0.00, 0.20 },

            new double[] { 2.00, 0.20 },
            new double[] { 1.75, 0.25 },
            new double[] { 1.50, 0.30 }, // Health (0,2) vs Resource (1), but with base scaling (1xR HP, 0.1xR RP
            new double[] { 1.25, 0.35 },
            new double[] { 1.00, 0.40 },

            new double[] { 0.20, 0.03 },
            new double[] { 0.15, 0.06 },
            new double[] { 0.10, 0.09 }, // Link (1) vs Movement (1)
            new double[] { 0.05, 0.12 }, 
            new double[] { 0.00, 0.15 }
        };

        string Icon_Affinity(int Index, bool Active, int Aff)
        {
            int Number = Index;
            if (Aff > 2)
            {
                Number = Index - 14;
            }
            else
            {
                Number = Index - (Aff - 1) * 7;
            }
            string Add = null;
            if (Active)
            {
                Add = "A";
            }
            else
            {
                Add = "D";
            }
            switch(Index)
            {
                case 1: case 2: case 4: case 5: case 6:
                    return "Affinity-T1-" + Add + Number.ToString();
                case 8: case 9: case 11: case 12: case 13: case 15:
                    return "Affinity-T2-" + Add + Number.ToString();
                case 3: case 7: case 10: case 14:
                    return "Affinity-TX-" + Add + Number.ToString();
                case 16: case 17: case 18:
                    return "Affinity-D-" + Add + Number.ToString();
                default: // also for 19
                    return "Affinity-TX-" + Add + "3";
            }
        }

        string[] AffinityNames = new string[6] { "Offense Trio Affinity", "Defense Trio Affinity", "Strike Affinity", "Survival Affinity", "Endurance Affinity", "Flexibility Affinity" };
        string[] ARed = new string[6] { "Physical Prowess", "Armor and Avoidance", "All Power", "Armor, Warding and Ease", "Health Points", "Link Points" };
        string[] ABlue = new string[6] { "Magical Prowess", "Warding and Resistance", "Precision", "Avoidance, Resistance and Attune", "Resource Points", "Movement Points" };
        string[] AYellow = new string[4] { "Healing Prowess", "Ease and Attune", "", "Attune" };
        int[] AIndexBoost = new int[6] { -1, -1, 11, 6, 16, 21 };

        void TT_Affinity_Load()
        {
            string PbAffName = null;
            PictureBox PbA = null;

            string[] Values = new string[14]
            {
                string.Format("{0}",AffinityStats[4]),
                string.Format("{0} and {1}",AffinityStats[6], AffinityStats[7]),
                string.Format("{0}",AffinityStats[5]),
                string.Format("{0}, {1} and {2}",AffinityStats[6], AffinityStats[10], AffinityStats[14]),
                string.Format("{0}",AffinityStats[0]),
                string.Format("{0}",AffinityStats[2]),

                string.Format("{0}",AffinityStats[8]),
                string.Format("{0} and {1}",AffinityStats[10], AffinityStats[11]),
                string.Format("{0}",AffinityStats[16]),
                string.Format("{0}, {1} and {2}",AffinityStats[7], AffinityStats[11], AffinityStats[15]),
                string.Format("{0}",AffinityStats[1]),
                string.Format("{0}",AffinityStats[3]),

                string.Format("{0}",AffinityStats[12]),
                string.Format("{0} and {1}",AffinityStats[14], AffinityStats[15])
            };

            for (int i = 0; i <= 5; i++)
            {
                if (TT_Affinity_Check)
                { TT_Affinity[i].Dispose(); }
                TT_Affinity[i] = new ToolTip();

                PbAffName = "PbAff" + (i + 1).ToString();
                PbA = this.Controls.Find(PbAffName, true).FirstOrDefault() as PictureBox;

                double[] Tuple = Affinity_Results.ElementAt(AffinityChoice[i] + AIndexBoost[i]);

                TT_Affinity[i].ToolTipTitle = AffinityNames[i];
                string TT = string.Format("Your {0} choice gives you the following bonusses (rounded down).{1}{1}", AffinityNames[i], Environment.NewLine);
                TT += string.Format("Red Track: {1}.{0}{3}% of Rank {4}.{0}{2} total.{0}{0}", Environment.NewLine, ARed[i], Values[i], Tuple[0] * 100, Character_Data[1] );
                TT += string.Format("Blue Track: {1}.{0}{3}% of Rank {4}.{0}{2} total.{0}{0}", Environment.NewLine, ABlue[i], Values[i + 6], Tuple[1] * 100, Character_Data[1] );
                if (i < 2)
                {
                    TT += string.Format("Yellow Track: {1}.{0}{3}% of Rank {4}.{0}{2} total.{0}{0}", Environment.NewLine, AYellow[i], Values[i + 12], Tuple[2] * 100, Character_Data[1] );
                }
                if (i == 1 || i == 3)
                {
                    TT += string.Format("These values are the result of both the {0} and the {1}.", AffinityNames[1], AffinityNames[3]);
                }
                TT_Affinity[i].SetToolTip(PbA, TT);
            }
            TT_Affinity_Check = true;
        }

        //
        //          Changing Abilities (Tab 3)
        //

        private void Ability_Activate(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int NrAbi = int.Parse(TempBox.Name.Substring(9, 1)); // Ability Place
            int NrAbiS = 0; // Which spot
            try { NrAbiS = int.Parse(TempBox.Name.Substring(11, 2)); } // For the double digits
            catch { NrAbiS = int.Parse(TempBox.Name.Substring(11, 1)); } // If not double, then single digit

            Ability_Change(NrAbi, NrAbiS);
            Ability_Growth_Update();
            Ability_Visibility();
            Ability_Calculation(NrAbi - 1, Character_Data[0]);
            GC.Collect(); // To use all those picture shizzle
        }
        private void Ability_Change(int NrAbility, int NrAbilitySelect)
        {
            if (AbilityChoice[NrAbility - 1] != NrAbilitySelect)
            {
                for (int i = 0; i <= 10; i++) // There are 11 nodes, 0 to 10
                {
                    string PbAbiName = "PbAbility" + NrAbility.ToString() + "_" + i.ToString(); // PbAbility2_0 for example
                    PictureBox ThisBox = this.Controls.Find(PbAbiName, true).FirstOrDefault() as PictureBox;
                    int Number = 0; // Standard = Deactivated

                    if (i == 0 && NrAbilitySelect > 0)
                    {
                        Number = 2; // Start Node
                    }
                    else if (i == 0 && NrAbilitySelect == 0)
                    {
                        Number = 1; // Zero Node only
                    }
                    else if (i > 0 && i < NrAbilitySelect)
                    {
                        Number = 3; // Center Nodes
                    }
                    else if (i == NrAbilitySelect)
                    {
                        Number = 4; // End Node
                    }
                    ImageChanger(ThisBox, Icon_Ability(Number));

                }
                AbilityChoice[NrAbility - 1] = NrAbilitySelect;
            }

            // Update the label quote
            string LblAbiName = "LblAbilityL" + NrAbility.ToString();
            Label ThisLbl = this.Controls.Find(LblAbiName, true).FirstOrDefault() as Label;
            ThisLbl.Text = AbilityLevels[NrAbilitySelect];
        }
        void Ability_Growth_Update()
        {
            AbilityGrowthPoints = AbilityGrowthPointsMax;
            for (int i = 1; i <= 9; i++) // All 9 abilities
            { AbilityGrowthPoints -= AbilityChoice[i - 1]; }
            LblAbilityGrowth.Text = string.Format("{1}{0}of the{0}{2}{0}Ability{0}Growth{0}Points{0}Remaining", Environment.NewLine, AbilityGrowthPoints, AbilityGrowthPointsMax);
        }
        void Ability_Visibility()
        {
            for (int i = 1; i <= 9; i++)
            {
                int GrowthCheck = AbilityChoice[i - 1] + AbilityGrowthPoints;
                for (int j = 0; j <= 10; j++)
                {
                    string PbAbiName = "PbAbility" + i.ToString() + "_" + j.ToString();
                    PictureBox ThisBox = this.Controls.Find(PbAbiName, true).FirstOrDefault() as PictureBox;

                    if (j <= GrowthCheck)
                    {
                        ThisBox.Visible = true;
                    }
                    else
                    {
                        ThisBox.Visible = false;
                    }
                }
            }
        }
        string Icon_Ability(int Index)
        {
            string Base = "Ability-";
            switch(Index)
            {
                case 1:
                    return Base + "A-Zero";
                case 2:
                    return Base + "A-Start";
                case 3:
                    return Base + "A-Center";
                case 4:
                    return Base + "A-End";
                default:
                    return Base + "D";
            }

        }

        List<string> AbilityLevels = new List<string>
        {
            "Non-existant", // 0
            "Terrible",     // 1
            "Bad",          // 2
            "Mediocre",     // 3
            "Sufficient",   // 4
            "Decent",       // 5
            "Good",         // 6
            "Great",        // 7
            "Outstanding",  // 8
            "Amazing",      // 9
            "Perfect"       // 10
        };

        List<double> AbilityGrowthRate = new List<double>  // Everything starts at terrible, and you have 35 points to spend total
        {
            0,          // 0 (Non-existant)
            0.025,      // 1 (Terrible) ( = 1 / 40 levels)
            0.05,       // 2 (Bad) ( = 1 / 20 levels)
            0.075,      // 3 (Mediocre) ( = 1 / 13,3 levels)
            0.1,        // 4 (Sufficient) ( = 1 / 10 levels)
            0.125,      // 5 (Decent) ( = 1 / 8 levels)
            0.150,      // 6 (Good) ( = 1 / 6,6 levels)
            0.175,      // 7 (Great) ( = 1 / 5.7 levels)
            0.2,        // 8 (Outstanding) ( = 1 / 5 levels)
            0.225,      // 9 (Amazing) ( = 1 / 4,4 levels)
            0.25        // 10 (Perfect) ( = 1 / 4 levels)
        };

        void Ability_Calculation(int AbilityIndex, int RoleRank)
        {
            double Calc = ((double)RoleRank + TalentStats[28]) * AbilityGrowthRate.ElementAt(AbilityChoice[AbilityIndex]) * 0.5;
            AbilityStats[AbilityIndex] = Math.Floor(Calc);
            double TotalStats = Math.Floor(Calc) + GearStats[AbilityIndex + 19] + TalentStats[AbilityIndex + 19];
            string LblCharName = "LblChar" + (AbilityIndex + 23).ToString();
            Label ThisLbl = this.Controls.Find(LblCharName, true).FirstOrDefault() as Label;
            ThisLbl.Text = TotalStats.ToString();
        }

        void Ability_Calculate()
        {
            for (int i = 0; i <= 8; i++) // 23+
            {
                Ability_Calculation(i, Character_Data[0]);
            }
        }

        //
        //          Gear Loading (Tab 4)
        //
        string Icon_Gear(int Index)
        {
            string Base = "EI-I-";
            string Add = "";
            if (Index > 15) 
            { 
                Add = "2"; 
                Index -= 16; 
            }
            switch(Index)
            {
                case 1: // Head
                    return Base + "AHelm" + Add;
                case 2: // Shoulders
                    return Base + "AShoulders" + Add;
                case 3: // Cloak
                    return Base + "ACloak" + Add;
                case 4: // Chest
                    return Base + "AChest" + Add;
                case 5: // No Weapon
                    return Base + "WN" + Add;
                case 6: // Gloves
                    return Base + "AGloves" + Add;
                case 7: // Pants
                    return Base + "APants" + Add;
                case 8: // Boots
                    return Base + "ABoots" + Add;
                case 9: // Trinket - Medal
                    return Base + "TM" + Add;
                case 10: // Trinket - Ring
                    return Base + "T" + Add;
                case 11: // Trinket - Necklace
                    return Base + "TN" + Add;
                case 12: // Trinket - Orb
                    return Base + "TO" + Add;
                case 13: // Weapon - Martial
                    return Base + "WP" + Add;
                case 14: // Weapon - Magical
                    return Base + "WM" + Add;
                case 15: // Weapon - Supporting
                    return Base + "WH" + Add;
                default: // None
                    return Base + "N" + Add;
            }
        }
        string Icon_Stat(int Index)
        {
            string Base = "Stat-";
            if (Index > 21) { Base = "EI-AB" + (Index-21).ToString() + "-"; }
            switch (Index)
            {
                case 0:
                    return Base + "Comment";
                case 1:
                    return Base + "Source";
                case 2:
                    return Base + "Lvl";
                case 3:
                    return Base + "HP";
                case 4:
                    return Base + "RP";
                case 5:
                    return Base + "LP";
                case 6:
                    return Base + "MP";
                case 7:
                    return Base + "PProw";
                case 8:
                    return Base + "PPower";
                case 9:
                    return Base + "Armor";
                case 10:
                    return Base + "Avoid";
                case 11:
                    return Base + "MProw";
                case 12:
                    return Base + "MPower";
                case 13:
                    return Base + "Ward";
                case 14:
                    return Base + "Resist";
                case 15:
                    return Base + "HProw";
                case 16:
                    return Base + "HPower";
                case 17:
                    return Base + "Rgn";
                case 18:
                    return Base + "Attune";
                case 19:
                    return Base + "Prec";
                case 20:
                    return Base + "CritHit";
                case 21:
                    return Base + "CritFail";
                case 22:
                    return Base + "Strength";
                case 23:
                    return Base + "Control";
                case 24:
                    return Base + "Perception";
                case 25:
                    return Base + "Knowledge";
                case 26:
                    return Base + "Magica";
                case 27:
                    return Base + "Survival";
                case 28:
                    return Base + "Deception";
                case 29:
                    return Base + "Finesse";
                case 30:
                    return Base + "Mentality";
                default:
                    return Base + "Lvl";
            }
        }
        public static class Gear
        {
            public static string[] Name = new string[12];
            public static string[] Type = new string[12];
            public static string[] Source = new string[12];
            public static string[] Description = new string[12];
            public static bool[] Equiped = new bool[12]; // not in class
            public static string[] Quality = new string[12];
            public static List<int[]> Stats = new List<int[]>()
            { new int[31], new int[31], new int[31], new int[31], new int[31], new int[31], new int[31], new int[31], new int[31], new int[31], new int[31], new int[31] };

            public static bool TT_Check = false;

            // Temp-bucket for loading gear
            public static string TempName = "None";
            public static string TempType = "";
            public static string TempSource = "";
            public static string TempDescription = "";
            public static int[] TempStats = new int[31];

            public static List<string> StatNames = new List<string>()
            {
              "Item Code",          // 0
              "Slot",               // 1
              "Suggested Skill",    // 2
              "Health Points",      // 3
              "Resource Points",    // 4
              "Link Points",        // 5
              "Movement Points",    // 6
              "Physical Prowess",   // 7
              "Physical Power",     // 8
              "Armor",              // 9
              "Avoidance",          // 10
              "Magical Prowess",    // 11
              "Magical Power",      // 12
              "Warding",            // 13
              "Resistance",         // 14
              "Healing Prowess",    // 15
              "Healing Power",      // 16
              "Ease",               // 17
              "Attune",             // 18
              "Precision",          // 19
              "Critical Success",   // 20
              "Critical Failure",   // 21
              "Strength",           // 22
              "Control",            // 23
              "Perception",         // 24
              "Knowledge",          // 25
              "Magica",             // 26
              "Survival",           // 27
              "Deception",          // 28
              "Finesse",            // 29
              "Mentality"           // 30
            };

            public static List<double> Budget = new List<double>()
            {
                0.2,    // 0: HP
                1,      // 1: RP
                1,      // 2: LP
                1,      // 3: MP
                0.5,    // 4: PProw
                0.5,    // 5: PPow
                0.5,    // 6: Armor
                1,      // 7: Avoid
                0.5,    // 8: MProw
                0.5,    // 9: MPow
                0.5,    // 10: Ward
                1,      // 11: Res
                0.5,    // 12: HProw
                0.5,    // 13: HPow
                0.5,    // 14: Ease
                1,      // 15: Att
                1,      // 16: Prec
                2,      // 17: Crit+
                -2,     // 18: Crit-

                1,      // 19: Str
                1,      // 20: Con
                1,      // 21: Per

                1,      // 22: Kno
                1,      // 23: Mag
                1,      // 24: Sur

                1,      // 25: Dec
                1,      // 26: Fin
                1       // 27: Men
            };

            public static List<string> LevelDesc = new List<string>()
        {
            "Cursed ",       // 0
            "Trash ",        // 1 (=<1)
            "Starter ",      // 2
            "Well-Made ",    // 3
            "Apprentice ",   // 4
            "Journeyman ",   // 5
            "Exceptional ",  // 6
            "Mastercraft ",  // 7
            "Ascendant ",    // 8
            "Epic ",         // 9
            "Legendary "     // 10
        };
        }

        void TempGearLoad(int ItemCode)
        {
            int Skipper = 0;  // Skip a few rows that contain headers

            try
            {
                var GValues = GSheets.Service.Spreadsheets.Values.Get(GSheets.Code, GSheets.TabItem + GSheets.TabItemRange).Execute().Values;

                foreach (var GRow in GValues)
                {
                    Skipper++;

                    if (GRow[1] != null && Skipper > 3)
                    {
                        if (GRow[0].ToString() == ItemCode.ToString()) // If the index matches the selected index, we are golden
                        {
                            Gear.TempType = GRow[2].ToString();
                            Gear.TempName = GRow[3].ToString();
                            Gear.TempSource = GRow[5].ToString();
                            Gear.TempDescription = GRow[34].ToString();
                            int Index = 0;

                            for (int i = 0; i <= 34; i++)
                            {
                                if (i != 2 && i != 3 && i != 5 && i != 34)
                                {
                                    string Row = GRow[i].ToString();
                                    if (Row == "") { Row = "0"; }
                                    Gear.TempStats[Index] = int.Parse(Row);
                                    Index++;
                                }
                            }
                            break;
                        }
                    }
                }
                GValues.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Breaking the items" + Skipper.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void TempGearPrintOut()
        {
            int UsedBox = 2;
            string PbGearName = null;
            string LblGearName = null;
            PictureBox ThisBox = null;
            Label ThisLbl = null;

            for (int i = 1; i <= 12; i++) // Reset everything
            {
                PbGearName = "PbGearStats" + i.ToString();
                LblGearName = "LblGearStats" + i.ToString();
                ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
                ThisLbl = this.Controls.Find(LblGearName, true).FirstOrDefault() as Label;
                ThisBox.Image = null;
                ThisLbl.Text = "";
            }

            int ImageSlot = GearSlotImage(Gear.TempStats[1], Gear.TempType);
            ImageChanger(PbGearStats1, Icon_Gear(ImageSlot));
            LblGearStats1.Text = string.Format("{1}{0}", Gear.TempType, GearRater(Gear.TempStats));

            for (int i = 2; i <= 30; i++) // First without abilities (21)
            {
                if (Gear.TempStats[i] != 0 && UsedBox < 12)
                {
                    PbGearName = "PbGearStats" + UsedBox.ToString();
                    LblGearName = "LblGearStats" + UsedBox.ToString();
                    ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
                    ThisLbl = this.Controls.Find(LblGearName, true).FirstOrDefault() as Label;

                    ImageChanger(ThisBox, Icon_Stat(i));
                    ThisLbl.Text = string.Format("{0}:{1}", Gear.StatNames.ElementAt(i), Gear.TempStats[i]);

                    UsedBox++;
                }
            }
            if (Gear.TempSource != "" && UsedBox < 12)
            {
                PbGearName = "PbGearStats" + UsedBox.ToString();
                LblGearName = "LblGearStats" + UsedBox.ToString();
                ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
                ThisLbl = this.Controls.Find(LblGearName, true).FirstOrDefault() as Label;

                ImageChanger(ThisBox, Icon_Stat(1));
                ThisLbl.Text = Gear.TempSource;

                UsedBox++;
            }

            PbGearName = "PbGearStats" + UsedBox.ToString();
            LblGearName = "LblGearStats" + UsedBox.ToString();
            ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
            ThisLbl = this.Controls.Find(LblGearName, true).FirstOrDefault() as Label;

            ImageChanger(ThisBox, Icon_Stat(0));
            ThisLbl.Text = string.Format(Gear.TempDescription, System.Environment.NewLine);
        }
        void TempGearListReset()
        {
            try
            {
                var GValues = GSheets.Service.Spreadsheets.Values.Get(GSheets.Code, GSheets.TabItem + GSheets.TabItemRange).Execute().Values; // Values is made up of a list of rows with columns as index
                CbGearSelect.Items.Clear(); // Refresh this list
                CbGearSelectIndex.Clear();

                foreach (var GRow in GValues)
                {
                    try
                    {
                        int.Parse(GRow[1].ToString()); // Check if the Item Slot exists
                        int Code = int.Parse(GRow[0].ToString()); // Remember the code

                        if (GearSelectiveSlot == 0)
                        {
                            string ItemNameType = string.Format("{0}-{1}: {2}", GRow[0].ToString(), GRow[2].ToString(), GRow[3].ToString()); // Add the type in front for better navigating
                            CbGearSelect.Items.Add(ItemNameType); // Add the names to the combobox
                            CbGearSelectIndex.Add(Code);
                        }
                        else
                        {
                            string ItemName = string.Format("{0}-{1}", GRow[0].ToString(), GRow[3].ToString()); // Add only the code, type is not needed.
                            if (GRow[1].ToString() == GearSelectiveSlot.ToString())
                            { 
                                CbGearSelect.Items.Add(ItemName);
                                CbGearSelectIndex.Add(Code);                      
                            }
                        }
                    }
                    catch
                    { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Too Bad", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CbGearSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            TempGearLoad(CbGearSelectIndex.ElementAt(CbGearSelect.SelectedIndex));
            TempGearPrintOut();
            EquipVisible(Gear.TempStats[1]);
            GC.Collect(); // Collect some garbage
        }
        List<int> CbGearSelectIndex = new List<int>();
        private void GearSearch_Toggle(object sender, EventArgs e)
        {
            PbGearSearchX.Image.Dispose();
            if (GearSelectiveSearch)
            {
                PbGearSearchX.Image = Properties.Resources.Gear_SearchAny;
                GearSelectiveSearch = false;
                GearSelect_Toggle(GearSelectiveSlot);
            }
            else
            {
                PbGearSearchX.Image = Properties.Resources.Gear_SearchSelective;
                GearSelectiveSearch = true;
            }
            
            for (int i = 1; i <= 9; i++)
            {
                string PbGearSName = "PbGearSearch" + i.ToString();
                PictureBox ThisBox = this.Controls.Find(PbGearSName, true).FirstOrDefault() as PictureBox;
                ThisBox.Visible = GearSelectiveSearch;
            }
        }
        private void PbGearSelect_Click(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int Nr = 0;
            try
            { Nr = int.Parse(TempBox.Name.Substring(12, 2)); }
            catch
            { Nr = int.Parse(TempBox.Name.Substring(12, 1)); }

            GearSelect_Toggle(Nr);
        }
        void GearSelect_Toggle(int Slot)
        {
            string PbGearSName = "PbGearSearch" + Slot.ToString();
            PictureBox ThisBox = this.Controls.Find(PbGearSName, true).FirstOrDefault() as PictureBox;

            if (Slot == GearSelectiveSlot) // If it's the type already chosen
            {
                GearSelectiveSlot = 0; // No Type
                ThisBox.Image = Resources.Gear_SearchSelectTypeD;
            }
            else // If it's a new type
            {
                if (GearSelectiveSlot != 0)
                {
                    string PbGearSOldName = "PbGearSearch" + GearSelectiveSlot.ToString();
                    PictureBox ThatBox = this.Controls.Find(PbGearSOldName, true).FirstOrDefault() as PictureBox;
                    ThatBox.Image = Resources.Gear_SearchSelectTypeD;
                }
                GearSelectiveSlot = Slot;
                ThisBox.Image = Resources.Gear_SearchSelectTypeA;
            }
            TempGearListReset();
        }
        void EquipVisible(int Slot = 0) // No additions = Invisible
        {
            if (Slot != 0)
            {
                string PbGearName = null;
                PictureBox ThisBox = null;

                if (Slot < 9) // The Gear item is not a trinket
                {
                    ThisBox = this.Controls.Find("PbGearEquip1", true).FirstOrDefault() as PictureBox;
                    if (Gear.Equiped[Slot - 1])
                    {
                        ThisBox.Image = Properties.Resources.Gear_Replace;
                    }
                    else
                    {
                        ThisBox.Image = Properties.Resources.Gear_Equip;
                    }
                    ThisBox.Visible = true;

                    PbGearEquip9.Visible = false;
                    PbGearEquip10.Visible = false;
                    PbGearEquip11.Visible = false;
                    PbGearEquip12.Visible = false;
                }
                else // The Gear item is a trinket
                {
                    for (int i = 9; i <= 12; i++)
                    {
                        PbGearName = "PbGearEquip" + i.ToString();
                        ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
                        if (Gear.Equiped[i - 1])
                        {
                            ThisBox.Image = Properties.Resources.Gear_ReplaceMini;
                        }
                        else
                        {
                            ThisBox.Image = Properties.Resources.Gear_EquipMini;
                        }
                        ThisBox.Visible = true;
                    }
                    PbGearEquip1.Visible = false;
                }
            }
            else
            {
                PbGearEquip1.Visible = false;
                PbGearEquip9.Visible = false;
                PbGearEquip10.Visible = false;
                PbGearEquip11.Visible = false;
                PbGearEquip12.Visible = false;
            }
        }
        private void Equip_Click(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int Nr = 0;
            int Slot = 0;
            try
            { Nr = int.Parse(TempBox.Name.Substring(11, 2)); }
            catch
            { Nr = int.Parse(TempBox.Name.Substring(11, 1)); }
            if (Nr >= 9)
            { Slot = Nr; }
            else
            { Slot = Gear.TempStats[1]; }

            if (Gear.Equiped[Slot - 1])
            {
                DialogResult dialogResult = MessageBox.Show(string.Format("Do you want to replace{0}{1}{0}with{0}{2}?", Environment.NewLine, Gear.Name[Slot - 1], Gear.TempName), "Replace Item", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    GearEquip(Slot);
                    EquipVisible();
                    Gear_Calculate();
                    Character_Data_Print();
                    TT_Character_Load();
                }

            }
            else
            {
                GearEquip(Slot);
                EquipVisible();
                Gear_Calculate();
                Character_Data_Print();
                TT_Character_Load();
            }

            TT_Gear_Load();

        }
        private void Unequip_Click(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int Nr = 0;
            try
            { Nr = int.Parse(TempBox.Name.Substring(6, 2)); }
            catch
            { Nr = int.Parse(TempBox.Name.Substring(6, 1)); }

            if (Gear.Equiped[Nr - 1])
            {
                DialogResult dialogResult = MessageBox.Show(string.Format("Do you want to unequip {0}?", Gear.Name[Nr - 1]), "Unequip Item", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    GearUnequip(Nr);
                    Gear_Calculate();
                    Character_Data_Print();
                    TT_Character_Load();
                }
            }

            TT_Gear_Load();
        }
        int GearSlotImage(int Slot, string Type)
        {
            int ImageSlot = Slot;
            if (Slot == 5)
            {
                if (Type == "Martial Weapon")
                {
                    ImageSlot = 13;
                }
                else if (Type == "Magical Weapon")
                {
                    ImageSlot = 14;
                }
                else if (Type == "Supporting Weapon")
                {
                    ImageSlot = 15;
                }
                else // Fist
                {
                    ImageSlot = 5;
                }

            }
            else if (Slot >= 9)
            {
                if (Type == "Medal")
                {
                    ImageSlot = 9;
                }
                else if (Type == "Ring")
                {
                    ImageSlot = 10;
                }
                else if (Type == "Necklace")
                {
                    ImageSlot = 11;
                }
                else // Orb
                {
                    ImageSlot = 12;
                }
            }
            return ImageSlot;
        }
        void GearEquip(int Slot, int ItemCode = 0)
        {
            if (ItemCode > 0)
            {
                TempGearLoad(ItemCode);
            }

            for (int i = 0; i < 31; i++)
            {
                Gear.Stats[Slot - 1][i] = Gear.TempStats[i]; // Convert the temp item to a real slot
            }

            Gear.Name[Slot - 1] = Gear.TempName;
            Gear.Type[Slot - 1] = Gear.TempType;
            Gear.Source[Slot - 1] = Gear.TempSource;
            Gear.Description[Slot - 1] = Gear.TempDescription;
            Gear.Equiped[Slot - 1] = true;
            Gear.Quality[Slot - 1] = GearRater(Gear.Stats[Slot - 1]);
            string PbGearName = "PbGear" + Slot.ToString();
            PictureBox ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
            ImageChanger(ThisBox, Icon_Gear(GearSlotImage(Slot, Gear.Type[Slot - 1])));
        }
        void GearUnequip(int Slot)
        {
            Array.Clear(Gear.Stats[Slot - 1], 0, Gear.Stats[Slot - 1].Length);
            Gear.Name[Slot - 1] = "";
            Gear.Type[Slot - 1] = "";
            Gear.Source[Slot - 1] = "";
            Gear.Description[Slot - 1] = "";
            Gear.Equiped[Slot - 1] = false;
            string PbGearName = "PbGear" + Slot.ToString();
            PictureBox ThisBox = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;
            ImageChanger(ThisBox, Icon_Gear(Slot + 16));
        }
        void TheNakedTruth()
        {
            for (int i = 1; i <= 12; i++)
            {
                GearUnequip(i);
            }
            TT_Gear_Load();
        }
        string GearRater(int[] Stats)
        {
            double TotalBudget = 1;
            bool Cursed = false;
            string Comment = "";
            for (int i = 0; i <= 27; i++)
            {
                if (Stats[i] < 0) { Cursed = true; }
                TotalBudget += (double)Stats[i + 3] * Gear.Budget[i];
            }
            if (TotalBudget < 0) { TotalBudget = 0; } else if (TotalBudget > 10) { TotalBudget = 10; }
            TotalBudget = Math.Floor(TotalBudget);

            if (Cursed)
            { Comment = string.Format("Cursed {0}", Gear.LevelDesc[(int)TotalBudget]); }
            else
            { Comment = Gear.LevelDesc[(int)TotalBudget]; }

            return Comment;
        }
        void Gear_Calculate()
        {
            Array.Clear(GearStats, 0, GearStats.Length);

            for (int i = 0; i <= 11; i++)
            {
                for (int j = 0; j < 27; j++)
                {
                    GearStats[j] += Gear.Stats[i][j + 3];
                }
            }
        }
        void TT_Gear_Load()
        {
            string PbGearName = null;
            PictureBox PbG = null;

            for (int i = 0; i < 12; i++)
            {
                if (Gear.TT_Check)
                { TT_Gear[i].Dispose(); } // Always refresh
                TT_Gear[i] = new ToolTip();

                PbGearName = "PbGear" + (i + 1).ToString();
                PbG = this.Controls.Find(PbGearName, true).FirstOrDefault() as PictureBox;

                if (Gear.Name[i] != null && Gear.Name[i] != "")
                {
                    TT_Gear[i].ToolTipTitle = Gear.Name[i];
                    string TT = string.Format("{0}{1}", Gear.Quality[i], Gear.Type[i]);
                    if (Gear.Source[i] != null && Gear.Source[i] != "")
                    { TT += string.Format("{0}Source: {1}", Environment.NewLine, Gear.Source[i]); }

                    for (int j = 2; j <= 30; j++)
                    {
                        if (Gear.Stats[i][j] != 0)
                        {
                            TT += string.Format("{0}{2} {1}", Environment.NewLine, Gear.StatNames[j], Gear.Stats[i][j]);
                        }
                    }
                    TT += string.Format("{0}" + Gear.Description[i], Environment.NewLine); // Weird structure to use {0} from Sheets

                    TT_Gear[i].SetToolTip(PbG, TT);
                }
                else
                {
                    TT_Gear[i].ToolTipTitle = "Nothing Equiped";
                    TT_Gear[i].SetToolTip(PbG, "Equip an item from the menu.");
                }
            }

            Gear.TT_Check = true;
        }

        //
        //          Becoming Talented (Tab 5)
        //       
        public class Talent
        {
            public int Cost = 0;
            public int Exclusive = 0;
            public string Name = null;
            public string Description = null;
            public string Effect = null;
            public ToolTip TT = new ToolTip() 
            {
                BackColor = SColor.Black, 
                ForeColor = SColor.White,
            };

            public void TT_Generate(PictureBox Box)
            {
                TT.ToolTipTitle = Name;
                string Caption = string.Format("Point Cost: {1}{0}" + Description + "{0}" + Effect, Environment.NewLine, Cost);
                TT.SetToolTip(Box, Caption);
            }
        }

        List<Talent> Talents = new List<Talent>()
        {
            // AFFINITY TALENTS
            new Talent
            {
                Cost = 15,
                Name = "Natural Affinity: Strike",
                Description = "You've got this!{0}Your Strike affinity will have a higher maximum to reach!",
                Effect = "(Get 1 additional growth rank to the chosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 15,
                Name = "Natural Affinity: Survival",
                Description = "You've got this!{0}Your Survival affinity will have a higher maximum to reach!",
                Effect = "(Get 1 additional growth rank to the chosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 15,
                Name = "Natural Affinity: Endurance",
                Description = "You've got this!{0}Your Endurance affinity will have a higher maximum to reach!",
                Effect = "(Get 1 additional growth rank to the chosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 15,
                Name = "Natural Affinity: Flexibility",
                Description = "You've got this!{0}Your Flexibility affinity will have a higher maximum to reach!",
                Effect = "(Get 1 additional growth rank to the chosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 10,
                Name = "Hybrid Potency: Strike",
                Description = "I get it, you do not want downsides.{0}Your Strike affinity will have a higher minimum.{0}No dump stats for you, my friend.",
                Effect = "(Get 1 additional growth rank to the unchosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 10,
                Name = "Hybrid Potency: Survival",
                Description = "I get it, you do not want downsides.{0}Your Survival affinity will have a higher minimum.{0}No dump stats for you, my friend.",
                Effect = "(Get 1 additional growth rank to the unchosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 10,
                Name = "Hybrid Potency: Endurance",
                Description = "I get it, you do not want downsides.{0}Your Endurance affinity will have a higher minimum.{0}No dump stats for you, my friend.",
                Effect = "(Get 1 additional growth rank to the unchosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 10,
                Name = "Hybrid Potency: Flexibility",
                Description = "I get it, you do not want downsides.{0}Your Flexibility affinity will have a higher minimum.{0}No dump stats for you, my friend.",
                Effect = "(Get 1 additional growth rank to the unchosen part of the Affinity)"
            },
            new Talent
            {
                Cost = 12,
                Name = "Ironskin",
                Description = "They think they slash you.{0}They think wrong.",
                Effect = "(Get 1 additional growth rank to the Armor part of the Defense Affinity)"
            },
            new Talent
            {
                Cost = 12,
                Name = "Barrier",
                Description = "It might not be visible, but{0}it certainly works.",
                Effect = "(Get 1 additional growth rank to the Warding part of the Defense Affinity)"
            },
            new Talent
            {
                Cost = 12,
                Name = "Regenerate",
                Description = "It is very important to keep yourself alive.{0}And to do so, it's better to be very healable.",
                Effect = "(Get 1 additional growth rank to the Ease part of the Defense Affinity)"
            },
            // GEAR TALENTS
            new Talent
            {
                Cost = 26,
                Name = "Old Favorite",
                Description = "Everyone likes to smack others with something,{0}but nothing is as satisfying as using your favorite smacking stick.{0}Of course our competence gives prowess!",
                Effect = "(+3 Physical Prowess when using a martial weapon)"
            },
            new Talent
            {
                Cost = 26,
                Name = "Competent Wizardry",
                Description = "My weapon will do pew pew, kablamo, pazazzz!{0}Some more prowess will never hurt when using my magical stick",
                Effect = "(+3 Magical Prowess when using a magical weapon)"
            },
            new Talent
            {
                Cost = 26,
                Name = "My Faithfull Tools",
                Description = "Faith is our strength.{0}Using it wisely will allow us{0}When wearing the right tool{0}It shall give us more prowess",
                Effect = "(+3 Healing Prowess when using a supporting weapon)"
            },
            new Talent
            {
                Cost = 22,
                Name = "Anything Will Do",
                Description = "Just grab a stick or something.{0}Just not one of those conventional weapons.",
                Effect = "(+2 on all Prowess when not using a specified weapon)"
            },
            new Talent
            {
                Cost = 20,
                Name = "Comfortable Wear",
                Description = "This outfit fits like a charm!{0} You get 1 avoidance for every of the following{0}slots which has an item equiped:{0}Shoulders, Chest, Gloves, Pants",
                Effect = "(Get +1 Avoidance for each of these armor slots that you filled:{0}Shoulders, Chest, Gloves, Pants)"
            },
            new Talent
            {
                Cost = 20,
                Name = "Grounded in Reality",
                Description = "The power to resist{0} You get 1 resistance for every of the following{0}slots which has an item equiped:{0}Head, Cloak, Chest, Boots",
                Effect = "(Get +1 Resistance for each of these armor slots that you filled:{0}Head, Cloak, Chest, Boots)"
            },
            new Talent
            {
                Cost = 20,
                Name = "Healing Magnets",
                Description = "Use your trinkets to find solace{0} You get 1 attune for every of trinket{0}slot which has an item equiped",
                Effect = "(Get +1 Attune for each of these trinket armor slots that you filled)"
            },
            new Talent
            {
                Cost = 28,
                Name = "I Am Unbreakable",
                Description = "Who needs clothes anyway?{0}You get 1 armor, warding and ease if you aren't fully equiped.",
                Effect = "(Get +1 Armor, Warding and Ease){0}(Requires an unequiped Gear Slot)"
            },
            new Talent
            {
                Cost = 25,
                Name = "The One Ring",
                Description = "You only ever need one ring",
                Effect = "(Get +2 Resource Points){0}(Required exactly 1 Ring Trinket to be equiped)"
            },
            new Talent
            {
                Cost = 14,
                Name = "Medal of Honor",
                Description = "You served the people well{0}Wear your medals with pride!{0}If you have 4 medals equiped, get an{0}Abiltiy bonus!",
                Effect = "(Get +1 on all Abilities){0}(Required 4 Medal Trinkets to be equiped)"
            },
            new Talent
            {
                Cost = -20,
                Name = "Cursebearer",
                Description = "Don't you just love gear{0}with a fair bit of extra challenge?{0}Cursed Gear gives you additional penalties.",
                Effect = "(Get -2 Health Points for each Cursed Gear item equiped)"
            },
            // TRADE OFF TALENTS
            new Talent
            {
                Exclusive = 23,
                Cost = 12,
                Name = "Cowardice",
                Description = "Because why would you even try to hit something.{0}As long as you don't get hit, you do not lose.",
                Effect = "(Get -10% Precision Growth){0}(Get +10% Avoidance and Resistance Growth)"
            },
            new Talent
            {
                Exclusive = 22,
                Cost = 12,
                Name = "Recklessness",
                Description = "True fighters handle just take attacks,{0}and they will make sure to return the favor.",
                Effect = "(Get +10% Precision Growth){0}(Get -10% Avoidance and Resistance Growth)"
            },
            new Talent
            {
                Exclusive = 25,
                Cost = 12,
                Name = "Plaguedoctor",
                Description = "A little bit of sickness gives you so much more profit.{0}Becoming tougher to heal is well worth the better attunement!",
                Effect = "(Get -5 Ease){0}(Get +3 Attune)"
            },
            new Talent
            {
                Exclusive = 24,
                Cost = 12,
                Name = "Nine Lives",
                Description = "So petable, so hard to be pet.{0}So healable, so hard to be healed",
                Effect = "(Get +5 Ease){0}(Get -3 Attune)"
            },
            new Talent
            {
                Exclusive = 27,
                Cost = 12,
                Name = "Make it Count",
                Description = "Power is for the taking!{0}However, you will not be as resourceful as you could be.",
                Effect = "(Get -10% Resource Point Growth){0}(Get +10% Power Growth)"
            },
            new Talent
            {
                Exclusive = 26,
                Cost = 12,
                Name = "More, Not Stronger",
                Description = "The winner is the one who{0}can still keep going!",
                Effect = "(Get +10% Resource Point Growth){0}(Get -10% Power Growth)"
            },
            new Talent
            {
                Exclusive = 29,
                Cost = 16,
                Name = "Lone Wolf",
                Description = "Others only get in the way.{0}All this connected nonsense is nothing for you.{0}All your Link Points will become extra Resource Points instead.",
                Effect = "(Lose all natural Link Points, Gain them as Resource Points)"
            },
            new Talent
            {
                Exclusive = 28,
                Cost = 16,
                Name = "Connected",
                Description = "My friends are my power, literally!{0}As long as your pals know some decent moves..{0}All your Resource Points will become extra Link Points instead.",
                Effect = "(Lose all natural Resource Points, Gain them as Resource Points)"
            },
            new Talent
            {
                Cost = 16,
                Name = "Eye for an Eye",
                Description = "Eye for an eye.{0}Smackdown for a smackdown.",
                Effect = "(Lose all natural Armor, Gain it as Physical Power)"
            },
            new Talent
            {
                Cost = 16,
                Name = "Glass Cannon",
                Description = "Sacrifice all your defenses.{0}Obtain even more offenses.{0}Glass cannons do shoot hard!",
                Effect = "(Lose all natural Warding, Gain it as Magical Power)"
            },
            new Talent
            {
                Cost = 16,
                Name = "Selfless",
                Description = "You better dodge, because you{0}will barely get healed anymore.",
                Effect = "(Lose all natural Ease, Gain it as Healing Power)"
            },
            // GROWTH TALENTS
            new Talent
            {
                Cost = 24,
                Name = "Unlimited Power!",
                Description = "Unlimited might be an overstatement,{0}but every little bit of resource helps, right?",
                Effect = "(Get +10% Resource Point Growth)"
            },
            new Talent
            {
                Exclusive = 35,
                Cost = 24,
                Name = "Like the Wind",
                Description = "Sting like a butterfly, strike like the...{0}Wait...{0}No?{0}Well, GOTTA GO FAST THEN!",
                Effect = "(Get +10% Movement Point Growth)"
            },
            new Talent
            {
                Exclusive = 34,
                Cost = -24,
                Name = "Slow and Steady",
                Description = "Just taunt the enemy enough to come to you!",
                Effect = "(Get -10% Movement Point Growth)"
            },
            new Talent
            {
                Cost = 18,
                Name = "Trust me, I'm a Doctor",
                Description = "Wait, are you really a doctor?{0}'I can do the healing'{0}I do not feel safe{0}'But with this talent it's even more healing!",
                Effect = "(Get +3 Healing Power)"
            },
            new Talent
            {
                Exclusive = 38,
                Cost = 24,
                Name = "I Just Wanna Live",
                Description = "As long as you have health, you are not dead.{0}So better get more of it then!",
                Effect = "(Get +30% Health Point Growth)"
            },
            new Talent
            {
                Exclusive = 37,
                Cost = -24,
                Name = "Just Don't Get Hit",
                Description = "Pain and misery always hit the spot.{0}Knowing you can't lose what you haven't got.",
                Effect = "(Get -30% Health Point Growth)"
            },
            new Talent
            {
                Exclusive = 40,
                Cost = 20,
                Name = "High Roller",
                Description = "Just what if everything had more{0}chances to be the best?",
                Effect = "(Get +1 Critical Success)"
            },
            new Talent
            {
                Exclusive = 39,
                Cost = -20,
                Name = "Skillfull Mistake",
                Description = "We all like to live a little...{0}..but this just seems like asking for it.",
                Effect = "(Get +1 Critical Failure)"
            },
            new Talent
            {
                Cost = 14,
                Name = "Elemental Prowess",
                Description = "Your own element is something you should be good at.{0}Some extra prowess never hurts then!",
                Effect = "(+2 Prowess on all Origin skills)"
            },
            new Talent
            {
                Cost = 16,
                Name = "Source Overdrive",
                Description = "Your source of power is your core.{0}You should figure out it's full potential.",
                Effect = "(Get a complete elemental source)"
            },
            new Talent
            {
                Cost = 14,
                Name = "Connected Prowess",
                Description = "You should be a true teamplayer.{0}Or you are just using your friends for power..{0}Either way, prowess on link skills helps!",
                Effect = "(+2 Prowess on all Link skills)"
            },
            // SKILL TALENTS
            new Talent
            {
                Cost = 16,
                Name = "God Eater",
                Description = "Sometimes you make stupid decisions.{0}Sometimes it was intentional.{0}Either way, to quickly solve this problem, you need more prowess!",
                Effect = "(Get +6 on all Prowess){0}(Requires a (too) high level opponent)"
            },
            new Talent
            {
                Cost = 24,
                Name = "Combat Meditation",
                Description = "Being passive might save you one day.{0}You get 1 resource point at the end of your turn if you did not use your Main action to attack",
                Effect = "(Gain the 'Combat Meditation' Passive Skill)"
            },
            new Talent
            {
                Cost = 15,
                Name = "Initial T",
                Description = "It's gonna be so exciting!{0}On the first turn of combat, get double the movement points{0}and two Main actions. You do not have a Minor action.",
                Effect = "(Gain the 'Initial T' Passive Skill)"
            },
            new Talent
            {
                Cost = 20,
                Name = "Hyperfocus",
                Description = "Take a breather for once, and focus.{0}You'll notice that things aren't as hard as they seem.{0}You can use this stored power when you really need it.",
                Effect = "(Gain the 'Hyperfocus' Skill Set)"
            },
            new Talent
            {
                Cost = 25,
                Name = "Minor Bargain",
                Description = "More is more, and time is resource!{0}You can use Minor action indefinitly, but each time{0}you use it in a turn, it costs 1 extra resource point.",
                Effect = "(Gain the 'Minor Bargain' Passive Skill)"
            },
            new Talent
            {
                Exclusive = 50,
                Cost = 15,
                Name = "Experienced Adventurer",
                Description = "I get knocked down,{0}I get up again.{0}You are never gonna keep me down",
                Effect = "(Gain the 'Experienced Adventurer' Passive Skill)"
            },
            new Talent
            {
                Exclusive = 49,
                Cost = -15,
                Name = "Too Much Pain",
                Description = "I get knocked down,{0}I give up again.{0}You are always gonna keep me down",
                Effect = "(Gain the 'Too Much Pain' Passive Skill)"
            },
            new Talent
            {
                Cost = -18,
                Name = "Blind Master",
                Description = "I see in other ways..{0}..just not as good as I wanted..",
                Effect = "(Gain the 'Blind Master' Passive Skill)"
            },
            new Talent
            {
                Cost = 44,
                Name = "Self-inserted Heroics",
                Description = "You are always there!{0}When a party member gains skills, you can say{0}'I was also there'{0}Your party might start to hate you for this",
                Effect = "(Gain +10 Skill Ranks for all calculations)"
            },
            new Talent
            {
                Cost = 1,
                Name = "Yee of Yellow Faith",
                Description = "Once you get yellow, you always want to go back!{0}All Critical Hits are Unavoidable and Piercing.{0}All Critical Failures hit yourself.{0}Live a litte, become one with the yellow.",
                Effect = "(Gain the 'Yellow Faith' Passive Skill)"
            },
            new Talent
            {
                Cost = 28,
                Name = "Master of Many Faces",
                Description = "Everybody knows your name and fame.{0}But not your other name.{0}Or the other one.",
                Effect = "(Gain +10 Role Ranks for all calculations)"
            },
            // ABILITY TALENT
            new Talent
            {
                Cost = 10,
                Name = "True Ability",
                Description = "Reflecting on your abilities makes{0}them improve ever so slowly..{0}",
                Effect = "(Gain 1 Ability Growth Point){0}{0}This Talent can be chosen multiple times.{0}Select it with the left mousebutton,{0}and deselect it with the right one"
            }
        };

        int[] Talent_Box(int Index)
        {
            int[] V = new int[2];
            V[0] = (int)Math.Floor((double)Index / 11) + 1;
            V[1] = Index % 11; // j value

            return V;
        } // Which PictureBox is used for the Index?
        int Talent_Select_Index(int NrTal, int NrTalS)
        {
            int S = (NrTal - 1) * 11 + NrTalS;
            return S;
        } // What Index is used for the PictureBox?

        private void Talent_Activate(object sender, EventArgs e)
        {
            PictureBox TempBox = (PictureBox)sender;
            int NrTal = int.Parse(TempBox.Name.Substring(8, 1)); // Talent Place
            int NrTalS = 0; // Which Spot
            try { NrTalS = int.Parse(TempBox.Name.Substring(10, 2)); } // For the double digits
            catch { NrTalS = int.Parse(TempBox.Name.Substring(10, 1)); } // If not double, then single digit

            Talent_Change(NrTal, NrTalS);
            Talent_Label_Update();
            Talent_Calculate();
            GC.Collect(); // Collect all garbage from the picturebox idiotics
        }

        private void Talent_Activate_Ability(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && TalentPoints >= Talents[55].Cost) // Activate
            {
                Talent_Selected_Ability(true);
            }
            else if (e.Button == MouseButtons.Left && TalentPoints < Talents[55].Cost)
            {
                MessageBox.Show(string.Format("You have {2} Talent Points, but require {1} to obtain this Talent.{0}Remove other Talents or gain more points to get it.", Environment.NewLine, Talents[55].Cost, TalentPoints), "Not Enough Points", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
            else if (e.Button == MouseButtons.Right && TalentChoice[55] > 0 && AbilityGrowthPoints > 0) //  Deactivate
            {
                Talent_Selected_Ability(false);
            }
            else if (e.Button == MouseButtons.Right && TalentChoice[55] > 0 && AbilityGrowthPoints == 0)
            {
                MessageBox.Show(string.Format("You used all your additional Ability Growth Points.{0}To remove this talent, first remove one Ability Growth Point on the Ability Tab.", Environment.NewLine));
            }

            Ability_Growth_Update();
            Ability_Visibility();
            Talent_Label_Update();
        }

        private void Talent_Change(int NrTalent, int NrTalentS)
        {
            int Selection = Talent_Select_Index(NrTalent, NrTalentS);

            int Exclusive = Talents[Selection].Exclusive;
            if (TalentChoice[Exclusive] > 0 && Exclusive != 0)
            {
                MessageBox.Show(string.Format("You have already chosen '{0}', which is exclusive with this talent.{1}Deactivate '{0}' to choose this Talent.", Talents[Exclusive].Name, Environment.NewLine), "This Talent is Exclusive", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (TalentChoice[Selection] > 0) // Unpick the Talent
                {
                    Talent_Selected(NrTalent, NrTalentS, Selection, false);
                }
                else if (TalentPoints >= Talents[Selection].Cost) // Pick the Talent
                {
                    Talent_Selected(NrTalent, NrTalentS, Selection, true);
                }
                else
                {
                    MessageBox.Show(string.Format("You have {2} Talent Points, but require {1} to obtain this Talent.{0}Remove other Talents or gain more points to get it.", Environment.NewLine, Talents[Selection].Cost, TalentPoints), "Not Enough Points", MessageBoxButtons.OK, MessageBoxIcon.Hand);

                }
            }
        }

        void Talent_Reset()
        {
            for (int i = 1; i <= 5; i++)
            {
                for (int j = 0; j <= 10; j++)
                {
                    Talent_Selected(i, j, Talent_Select_Index(i, j), false);
                }
            }
            // Talent_Ability
            Talent_Point_Calculate();
        }

        void Talent_Selected_Ability(bool Activate, bool Pay = true)
        {
            if (Activate && Pay)
            {
                TalentChoice[55] += 1;
                AbilityGrowthPointsMax += 1;
                TalentPoints -= Talents[55].Cost;
            }
            else if (Pay)
            {
                TalentChoice[55] -= 1;
                AbilityGrowthPointsMax -= 1;
                TalentPoints += Talents[55].Cost;
            }
            else
            {
                TalentChoice[55] = 0;
            }
            if (TalentChoice[55] > 0)
            {
                ImageChanger(PbTalent6, "EI_T_55A");
                //PbTalent6.Image = Talents[55].Image_A;
            }
            else
            {
                ImageChanger(PbTalent6, "EI_T_55D");
                //PbTalent6.Image = Talents[55].Image_D;
            }
            LblTalent6.Text = TalentChoice[55].ToString();
        }

        void Talent_Selected(int NrT, int NrTS, int Select, bool Activate, bool Pay = true)
        {
            string PbTalName = "PbTalent" + NrT.ToString() + "_" + NrTS.ToString();
            PictureBox ThisBox = this.Controls.Find(PbTalName, true).FirstOrDefault() as PictureBox;

            if (Activate && Pay)
            {
                ImageChanger(ThisBox, "EI_T_" + Select.ToString() + "A");
                TalentChoice[Select] = 1;
                TalentPoints -= Talents[Select].Cost;
            }
            else if (Pay)
            {
                ImageChanger(ThisBox, "EI_T_" + Select.ToString() + "D");
                TalentChoice[Select] = 0;
                TalentPoints += Talents[Select].Cost;
            }
            else
            {
                ImageChanger(ThisBox, "EI_T_" + Select.ToString() + "D");
                TalentChoice[Select] = 0;
            }
        }

        void Talent_Point_Calculate()
        {
            TalentPointsMax = Character_Data[2] * 3;
            TalentPoints = TalentPointsMax;
            for (int i = 0; i <= 55; i++)
            {
                TalentPoints -= TalentChoice[i] * Talents[i].Cost;
            }
        }

        void Talent_Plot_Affinity()
        {
            // Show the Natural Affinity and Hybrid Potency signs, Talents are handled elsewhere
            PictureBox PbAff = null;
            string[] Talent_Affinity_String = new string[8] { "PbAff3_7", "PbAff4_7", "PbAff5_7", "PbAff6_7", "PbAff3_6", "PbAff4_6", "PbAff5_6", "PbAff6_6" };

            for (int i = 0; i < 8; i++)
            {
                PbAff = this.Controls.Find(Talent_Affinity_String[i], true).FirstOrDefault() as PictureBox;
                if (TalentChoice[i] > 0)
                {
                    PbAff.Visible = true;
                }
                else
                {
                    PbAff.Visible = false;
                }
            }
        }

        void Talent_Calculate()
        {
            Talent_Plot_Affinity();

            Array.Clear(TalentStats, 0, TalentStats.Length);
            if (TalentChoice[52] > 0) // Heroics
            {
                TalentStats[29] += 10;
            }
            if (TalentChoice[54] > 0) // Faces
            {
                TalentStats[28] += 10;
            }

            double SkillRank = (double)Character_Data[1] + TalentStats[29];
            double TenPercent = Math.Floor(SkillRank * 0.1);

            // Gear Talents
            if (TalentChoice[11] > 0 && Gear.Type[4] == "Martial Weapon")
            {
                TalentStats[4] += 3; // Ph Prow
            }
            if (TalentChoice[12] > 0 && Gear.Type[4] == "Magical Weapon")
            {
                TalentStats[8] += 3; // Ma Prow
            }
            if (TalentChoice[13] > 0 && Gear.Type[4] == "Supporting Weapon")
            {
                TalentStats[12] += 3; // He Prow
            }
            if (TalentChoice[14] > 0 && Gear.Type[4] != "Martial Weapon" && Gear.Type[4] != "Magical Weapon" && Gear.Type[4] != "Supporting Weapon")
            {
                TalentStats[4] += 2;
                TalentStats[8] += 2;
                TalentStats[12] += 2;
            }
            if (TalentChoice[15] > 0) // Slots equiped: Shoulders, Chest, Gloves, Pants > +1 Avoid
            {
                if (Gear.Equiped[1]) { TalentStats[7] += 1; }
                if (Gear.Equiped[3]) { TalentStats[7] += 1; }
                if (Gear.Equiped[5]) { TalentStats[7] += 1; }
                if (Gear.Equiped[6]) { TalentStats[7] += 1; }
            }
            if (TalentChoice[16] > 0) // Slots equiped: Head, Cloak, Chest, Boots > +1 Resist
            {
                if (Gear.Equiped[0]) { TalentStats[11] += 1; }
                if (Gear.Equiped[2]) { TalentStats[11] += 1; }
                if (Gear.Equiped[3]) { TalentStats[11] += 1; }
                if (Gear.Equiped[7]) { TalentStats[11] += 1; }
            }
            if (TalentChoice[17] > 0) // Slots equiped: Trinket > +1 Attune
            {
                for (int i = 8; i <= 11; i++)
                {
                    if (Gear.Equiped[i]) { TalentStats[15] += 1; }
                }
            }
            if (TalentChoice[18] > 0) // If you have a slot unequiped
            {
                bool Unequiped = false;
                for (int i = 0; i <= 11; i++)
                {
                    if (!Gear.Equiped[i]) { Unequiped = true; }
                }
                if (Unequiped)
                {
                    TalentStats[6] += 1; // Armor
                    TalentStats[10] += 1; // Ward
                    TalentStats[14] += 1; // Ease
                }
            }
            if (TalentChoice[19] > 0)
            {
                bool OneRing = false;
                for (int i = 8; i <= 11; i++)
                {
                    if (Gear.Type[i] == "Ring" && OneRing) { OneRing = false; }
                    else if (Gear.Type[i] == "Ring") { OneRing = true; }
                    if (OneRing)
                    {
                        TalentStats[1] += 2; // Resource
                    }
                }
            }
            if (TalentChoice[20] > 0 && Gear.Type[8] == "Medal" && Gear.Type[9] == "Medal" && Gear.Type[10] == "Medal" && Gear.Type[11] == "Medal")
            {
                for (int i = 19; i <= 27; i++)
                {
                    TalentStats[i] += 1; // All abilities
                }
            }
            if (TalentChoice[21] > 0)
            {
                for (int i = 0; i <= 11; i++)
                {
                    if (Gear.Quality[i] != null)
                    {
                        if (Gear.Quality[i].Substring(0, 6) == "Cursed")
                        {
                            TalentStats[0] -= 2; // Health
                        }
                    }
                }
            }

            // Trade Off Talents
            if (TalentChoice[22] > 0)
            {
                TalentStats[7] += TenPercent;
                TalentStats[11] += TenPercent;
                TalentStats[16] -= TenPercent; // Prec
            }
            else if (TalentChoice[23] > 0)
            {
                TalentStats[7] -= TenPercent;
                TalentStats[11] -= TenPercent;
                TalentStats[16] += TenPercent;
            }
            if (TalentChoice[24] > 0)
            {
                TalentStats[14] -= 5; // Ease
                TalentStats[15] += 3; // Attune
            }
            else if (TalentChoice[25] > 0)
            {
                TalentStats[14] += 5;
                TalentStats[15] -= 3;
            }
            if (TalentChoice[27] > 0)
            {
                TalentStats[1] += TenPercent; // RP
                TalentStats[5] -= TenPercent; // Ph Pow
                TalentStats[9] -= TenPercent; // Ma Pow
                TalentStats[13] -= TenPercent; // He Pow
            }
            else if (TalentChoice[26] > 0)
            {
                TalentStats[1] -= TenPercent; // RP
                TalentStats[5] += TenPercent; // Ph Pow
                TalentStats[9] += TenPercent; // Ma Pow
                TalentStats[13] += TenPercent; // He Pow
            }

            // Growth Talents
            if (TalentChoice[33] > 0)
            {
                TalentStats[1] += TenPercent; // RP
            }
            if (TalentChoice[34] > 0)
            {
                TalentStats[3] += TenPercent; // MP
            }
            if (TalentChoice[35] > 0)
            {
                TalentStats[3] -= TenPercent; // MP
            }
            if (TalentChoice[36] > 0)
            {
                TalentStats[13] += 3; // HPow
            }
            if (TalentChoice[37] > 0)
            {
                TalentStats[0] += 3 * TenPercent; // HP
            }
            if (TalentChoice[38] > 0)
            {
                TalentStats[0] -= 3 * TenPercent; // HP
            }
            if (TalentChoice[39] > 0)
            {
                TalentStats[17] += 1; // Crit Suc
            }
            if (TalentChoice[40] > 0)
            {
                TalentStats[18] += 1; // Crit Suc
            }
            // 41 & 43 are in Sheets
        }

        void Talent_Calculate_TradeOff()
        {
            double Trade = 0;

            if (TalentChoice[28] > 0) // LP > RP
            {
                Trade = AffinityStats[2] + TalentStats[2];
                TalentStats[2] -= Trade;
                TalentStats[1] += Trade;
            }
            else if (TalentChoice[29] > 0) // RP > LP
            {
                Trade = AffinityStats[1] + TalentStats[1];
                TalentStats[2] += Trade;
                TalentStats[1] -= Trade;
            }
            if (TalentChoice[30] > 0) // Arm > PPow
            {
                Trade = AffinityStats[6] + TalentStats[6];
                TalentStats[6] -= Trade;
                TalentStats[5] += Trade;
            }
            if (TalentChoice[31] > 0) // War > MPow
            {
                Trade = AffinityStats[10] + TalentStats[10];
                TalentStats[10] -= Trade;
                TalentStats[9] += Trade;
            }
            if (TalentChoice[32] > 0) // Eas > HPow
            {
                Trade = AffinityStats[14] + TalentStats[14];
                TalentStats[14] -= Trade;
                TalentStats[13] += Trade;
            }
        }

        void Talent_Label_Update()
        {
            LblTalentPoint.Text = string.Format("{1} Talent Points{0}of the{0}{2} available", Environment.NewLine, TalentPoints, TalentPointsMax);
        }

        void TT_Talent_Load()
        {
            int[] Nr = null;
            string PbTalName = null;

            for (int i = 0; i < Talents.Count; i++)
            {
                Nr = Talent_Box(i);
                if (i != 55) // There are no 6 subvalues at the moment
                { PbTalName = "PbTalent" + Nr[0].ToString() + "_" + Nr[1].ToString(); }
                else
                { PbTalName = "PbTalent" + Nr[0].ToString(); }

                Talents[i].TT_Generate(this.Controls.Find(PbTalName, true).FirstOrDefault() as PictureBox);
            }
        }

        /// 
        /// Source Elements
        /// 

        void SourceLoad(int Index)
        {
            PictureBox ThisBox = this.Controls.Find("PbChar3",true).FirstOrDefault() as PictureBox;
            Label ThisLbl = this.Controls.Find("LblChar3", true).FirstOrDefault() as Label;

            //ThisBox.Image = Source.Pic.ElementAt(Index); // IMAGECHANGER
            ImageChanger(ThisBox, Source.Icon_Source(Index));
            ThisLbl.Text = Source.Name(Index, TalentChoice[42] > 0);
        }

        private void CbCharSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            Character_Data[3] = CbCharSource.SelectedIndex;
            LblChar3.Text = CbCharSource.Text;

            SelectToggle(CbCharSource, LblChar3);
            SourceLoad(Character_Data[3]);
        }

        void SelectToggle(object A, object B)
        {
            ComboBox AC = null;
            TextBox AT = null;
            Label BL = (Label)B;

            if (A.GetType() == typeof(ComboBox))
            {
                AC = (ComboBox)A;

                if (AC.Visible)
                {
                    AC.Visible = false;
                    BL.Visible = true;
                }
                else
                {
                    AC.Visible = true;
                    BL.Visible = false;
                }
            }
            else if (A.GetType() == typeof(TextBox)) 
            {
                AT = ((TextBox)A);

                if (AT.Visible)
                {
                    AT.Visible = false;
                    BL.Visible = true;
                }
                else
                {
                    AT.Visible = true;
                    BL.Visible = false;
                }
            }
        }

        private void Source_Toggle(object sender, EventArgs e)
        {
            SelectToggle(CbCharSource, LblChar3);
        }

        private void Name_Toggle(object sender, EventArgs e)
        {
            SelectToggle(TxtCharName, LblCharName);
        }

        private void TxtCharName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Character_Name = TxtCharName.Text;
                LblCharName.Text = TxtCharName.Text;
                SelectToggle(TxtCharName, LblCharName);
            }
        }

        private void TxtRank_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TextBox TxtRank = (TextBox)sender;
                int NrRank = int.Parse(TxtRank.Name.Substring(7, 1));
                Label LblRank = this.Controls.Find("LblChar" + NrRank.ToString(), true).FirstOrDefault() as Label;
                LblRank.Text = TxtRank.Text;
                SelectToggle(TxtRank, LblRank);

                Character_Data[NrRank] = int.Parse(TxtRank.Text);
                Affinity_Calculate();
                Ability_Calculate();
                Talent_Calculate();
                Talent_Calculate_TradeOff();
                Talent_Point_Calculate(); // overwrites current points atm
                Talent_Label_Update();
                Character_Data_Return();
                Character_Data_Print();
                TT_Character_Load();

                // Recalc for new ranks
            }
        }

        private void Rank_Toggle(object sender, EventArgs e)
        {
            // string Name = (Label)sender.Name;
            Label LblRank = (Label)sender;
            int NrRank = int.Parse(LblRank.Name.Substring(7, 1));
            TextBox TxtRank = this.Controls.Find("TxtChar" + NrRank.ToString(), true).FirstOrDefault() as TextBox;
            TxtRank.Text = LblRank.Text;
            SelectToggle(TxtRank, LblRank);
        }

        private void PbExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PbHelp_Click(object sender, EventArgs e)
        {
            if (HelpMode)
            {
                MessageBox.Show(string.Format("This Tab will reflect all your character stats and abilities.{0}When you load a character, remake it, or create a new one and edit their specific traits on other tabs, this page will update. Double-click on the name, source or ranks to manually update those. This is not needed for current party members, as they will be automaticly updated! When you save a character, these calculated stats will be put into your trusty Sheets, updating all auto-calculation on your skills! Life was never this easy!", Environment.NewLine), "Support ALL the users!", MessageBoxButtons.OK, MessageBoxIcon.Question);

            }
            else
            {
                MessageBox.Show(string.Format("Help mode is now active, please look for the help icon on different tabs to get tips and tricks!{0}{0}This Tab will reflect all your character stats and abilities.{0}When you load a character, remake it, or create a new one and edit their specific traits on other tabs, this page will update. Double-click on the name, source or ranks to manually update those. This is not needed for current party members, as they will be automaticly updated! When you save a character, these calculated stats will be put into your trusty Sheets, updating all auto-calculation on your skills! Life was never this easy!", Environment.NewLine), "Support ALL the users!", MessageBoxButtons.OK, MessageBoxIcon.Question);
                HelpMode = true;
                PbHelp2.Visible = true;
                PbHelp3.Visible = true;
                PbHelp4.Visible = true;
                PbHelp5.Visible = true;
            }
        }

        private void PbHelp2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(string.Format("This Tab will show your Affinities.{0}Affinities give you certain choices that will affect the growth of your character. Affinity Growth is based on your Skill Rank, which you can gain in every game session.{0}Simply choose one of the hexagon markers and it will light up in the color of your choices. The icon will tell you what you are choosing. When you hover over the icon you will get more information on what you actually chose, and what the impact is at your current Skill Rank. If you are unsure, just pick the middle ground, and edit your affinity later when your character has been defined more!", Environment.NewLine), "Support ALL the users!", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void PbHelp3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(string.Format("This Tab will show your Abilities.{0}Abilities are mostly used outside of combat scenario's to define your success in a certain skill, and you choose your Ability Growth level. Ability Growth is based on your Role Rank, which you can gain in every game session.{0}Each ability has a starting value of 0, and goes up to 10. Pick the hexagon marker more to the right to get a higher Ability Growth Rate. Watch out though, because if you are very good in something, it also means being completely worthless in other fields. The growth will be reflected right away on your Character Tab. To get a slight edge on Ability Growth points, you can also get a Talent to increase your maximum!", Environment.NewLine), "Support ALL the users!", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void PbHelp4_Click(object sender, EventArgs e)
        {
            MessageBox.Show(string.Format("This Tab will show your Gear.{0}Gear can be acquired as treasure, from shops, crafters and more. This is the core progression besides ranks. You can use the drop down menu to look at items, then equip them by used the equip button. After you equip items, their powers will automaticly be used in the calculations for your Character tab. Hover over each gear icon to see what you have equiped and what that means! If you no longer have/want a piece of gear, simply click on it's icon and say 'yes' to the removal prompt.", Environment.NewLine), "Support ALL the users!", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void PbHelp5_Click(object sender, EventArgs e)
        {
            MessageBox.Show(string.Format("This Tab will show your Talents.{0}Talents are simular to Affinities and Abilities, but they are more widespread. You get Talent Points based on your Quest Rank, which you can gain in most game sessions.{0}Hovering over a talent shows it's description, power and cost. The icon can also show you a lot already.{0}{0}Affinity Talents are extra bonusses for your Affinities. Using them can further enhance your powers.{0}{0}Gear Talents give you perks when you have certain pieces of gear equiped. They can be quite strong, but also restricting.{0}{0}Trade Off Talents are quite cheap, but harsh. You might get a nice perk from them, but you also get negative value in return.{0}{0}Growth Talents are a bit more simple. They give you perks without much conditions. However, they are quite costly.{0}{0}Skill Talents are the odd ones out. They give you a new Skill on your character sheet, to be used in combat scenario's!{0}{0}As one can see talents have a bit of color coding and signals to them:{0}Faded Yellow is a basic Talent, giving perks.{0}Purple is a strained talent, giving talent points and a negative perk.{0}Yellow talents are skill talents{0}Red, blue, yellow talents are based on the affinity with that icon{0}The link lines on certain sets of talents mean they are exclusive; you can only pick one.", Environment.NewLine), "Support ALL the users!", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void PbSheets_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Do you want to refresh your connection to Google Sheets?", "Refresh Sheets", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    var GValues = GSheets.Service.Spreadsheets.Values.Get(GSheets.Code, GSheets.TabItem + GSheets.TabItemRange).Execute().Values; // Values is made up of a list of rows with columns as index
                    foreach (var GRow in GValues)
                    {
                        string Test = GRow[0].ToString();
                    }
                    MessageBox.Show("The load seems to work properly. Google Sheets should be connected!", "I'm not a virus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    MessageBox.Show("A browser window should have popped up. Make sure to press continue and accept the link! Once the browser tells you that the link has been established, close the browser and try again!", "I'm not a virus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("If you keep getting errors whilst typing a correct code, refreshing might be a good choice", "Next time maybe?", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    public static class Source
    {
        public static string Icon_Source(int Index)
        {
            int MainTrack = 0;
            if (Index%2 > 0) { MainTrack = (Index + 1) / 2; }
            else { MainTrack = Index / 2; }

            if (Index == 0)
            {
                return string.Format("S_{0}", MainName.ElementAt(0));
            }
            else
            {
                return string.Format("S_{0}_{1}", MainName.ElementAt(MainTrack), SubName.ElementAt(Index));
            }
        }

        public static string Name(int Index, bool Talent)
        {
            int MainTrack = 0;
            if (Index%2 > 0) { MainTrack = (Index + 1) / 2; }
            else { MainTrack = Index / 2; }

            if(Index == 0)
            {
                return MainName.ElementAt(0);
            }
            else if (Talent)
            {
                return string.Format("{0}: {1}", MainName.ElementAt(MainTrack), SubName.ElementAt(0));
            }
            else
            {
                return string.Format("{0}: {1}", MainName.ElementAt(MainTrack), SubName.ElementAt(Index));
            }
        }

        public static List<string> MainName = new List<string>()
        {
            "None", //0
            "Arcane", //1
            "Chaos", //2
            "Earth", //3
            "Energy", //4
            "Fire", //5
            "Lightning", //6
            "Light", //7
            "Nature", //8
            "Poison", //9
            "Primal", //10
            "Psi", //11
            "Shadow", //12
            "Void", //13
            "Water" //14
        };

        public static List<string> SubName = new List<string>()
        {
            "Prime", // 0
            "Illusion", // 1
            "Magic", // 2
            "Anomaly", // 3
            "Gambler", // 4
            "Mineral", // 5
            "Sand", // 6
            "Anima", // 7
            "Force", // 8
            "Flame", // 9
            "Mythical", // 10
            "Magnetism", // 11
            "Storm", // 12
            "Holy", // 13
            "Prismatic", // 14
            "Fae", // 15
            "Wind", // 16
            "Alchemic", // 17
            "Toxic", // 18
            "Beastial", // 19
            "Shamanistic", // 20
            "Ki", // 21
            "Psychic", // 22
            "Darkness", // 23
            "Necrotic", // 24
            "Consuming", // 25
            "Space", // 26
            "Flow", // 27
            "Ice"  // 28
        };
    }
}
