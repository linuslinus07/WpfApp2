using System.Text;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;
using System.Windows.Threading;
using System.Runtime.CompilerServices;
using System.Drawing;
using System.Runtime.Intrinsics.X86;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Net.WebRequestMethods;
using Microsoft.VisualBasic;
using System.Linq.Expressions;
using System.Windows.Media.Media3D;

namespace WpfApp2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += (s, e) => _time.Text = DateTime.Now.ToString("HH:mm:ss");
            timer.Start();

            #region -- Write Settings --
            _pathInput.Text = Properties.Settings.Default.pathInput.ToString();

            if (Properties.Settings.Default.IsChecked)
            {
                _height.Text = Properties.Settings.Default.height.ToString();
                _mat.Text = Properties.Settings.Default.mat.ToString();
                _diaInt.Text = Properties.Settings.Default.diaInt.ToString();
                _xErste.Text = Properties.Settings.Default.xErste.ToString();
                _rabo.Text = Properties.Settings.Default.rabo.ToString();
                _lab.Text = Properties.Settings.Default.lab.ToString();
                _xVersch.Text = Properties.Settings.Default.xVersch.ToString();
                _secur.Text = Properties.Settings.Default.secur.ToString();
                _durch.Text = Properties.Settings.Default.durch.ToString();
                _spinRPM.Value = Properties.Settings.Default.spinRPM;
                _deg.Text = Properties.Settings.Default.deg.ToString();
                _xAbst.Text = Properties.Settings.Default.xAbst.ToString();

                _speichern.IsChecked = true;

                string lastOption = Properties.Settings.Default.drillDia;

                switch (lastOption)
                {
                    case "0.8":
                        OptionA.IsChecked = true;
                        break;
                    case "1.0":
                        OptionB.IsChecked = true;
                        break;
                    case "1.3":
                        OptionC.IsChecked = true;
                        break;
                    case "1.5":
                        OptionD.IsChecked = true;
                        break;
                    case "2.0":
                        OptionE.IsChecked = true;
                        break;
                    case "3.0":
                        OptionF.IsChecked = true;
                        break;
                        // Add other cases as needed
                }
            }
            #endregion

        }

        private void ClearInputs(object sender, RoutedEventArgs e)
        {
            #region -- settings zurücksetzen --
            _diaInt.Text = "0";
            _height.Text = "0";
            _lab.Text = "4";
            _xErste.Text = "0";
            _deg.Text = "1";
            _mat.Text = "5";
            _xAbst.Text = "0";
            _rabo.Text = "0";
            _xVersch.Text = "0";
            _anz.Text = "1";
            _secur.Text = "8";
            _durch.Text = "2";
            _spinRPM.Value = 24000;

            // RadioButtons deaktivieren
            OptionA.IsChecked = false;
            OptionB.IsChecked = false;
            OptionC.IsChecked = false;
            OptionD.IsChecked = false;
            OptionE.IsChecked = false;
            OptionF.IsChecked = false;

            _speichern.IsChecked = false;

            // Settings zurücksetzen
            Properties.Settings.Default.height = 0;
            Properties.Settings.Default.lab = 4;
            Properties.Settings.Default.mat = 5;
            Properties.Settings.Default.drillDia = "0";
            Properties.Settings.Default.xErste = 0;
            Properties.Settings.Default.rabo = 0;
            Properties.Settings.Default.deg = 0;
            Properties.Settings.Default.durch = 2;
            Properties.Settings.Default.spinRPM = 24000;
            Properties.Settings.Default.xVersch = 0;
            Properties.Settings.Default.diaInt = 0;
            Properties.Settings.Default.secur = 8;
            Properties.Settings.Default.xAbst = 0;
            Properties.Settings.Default.IsChecked = false;
            Properties.Settings.Default.Save();
            #endregion
        }


        public async void Calculate(object sender, RoutedEventArgs e)
        {
            #region -- Calculate --
            try
            {
                #region -- Variabeln --
                if (!double.TryParse(_diaInt.Text, out double diaInt))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für den Innendurchmesser ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_height.Text, out double height))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für die Höhe ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_lab.Text, out double lab))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für den Lochabstand ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_xErste.Text, out double xErste))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für den Randabstand ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_deg.Text, out double deg))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für den Lochabstand Radial ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_mat.Text, out double mat))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für die Materialstärke ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_xAbst.Text, out double xAbst))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für den X-Abstand ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_rabo.Text, out double rabo))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für den Randabstand ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_xVersch.Text, out double xVersch))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für die X-Verschiebung ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_anz.Text, out double anz))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für die Anzahl ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_secur.Text, out double secur))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für die Sicherheitshöhe ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(_durch.Text, out double durch))
                {
                    MessageBox.Show("Bitte gib eine gültige Zahl für die Durchbruchtiefe ein.", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                double spinRPM = _spinRPM.Value;
                bool IsChecked = _speichern.IsChecked == true;
                //string directory = @"C:\Users\linus\Downloads\";
                string directory = _pathInput.Text;
                string fileName = $"d={diaInt}mm_h={height}mm.txt";
                string filePath = System.IO.Path.Combine(directory, fileName);

                double drillDia = OptionA.IsChecked == true ? 0.8 :
                          OptionB.IsChecked == true ? 1.0 :
                          OptionC.IsChecked == true ? 1.3 :
                          OptionD.IsChecked == true ? 1.5 :
                          OptionE.IsChecked == true ? 2.0 :
                          OptionF.IsChecked == true ? 3.0 : 0.0;

                double Currentx;
                double xPos;
                double aPos;
                int h;
                int i;
                bool Spind1;
                bool Spind2;
                bool Spind3;
                bool Spind4;

                double diaExt = diaInt + 2 * mat;
                double radExt = diaExt / 2;
                double zStart = radExt + secur;
                double zEnd = diaInt / 2 - durch;

                double nettoBohr = height - rabo - xErste;

                double labOffset = lab / 2;
                double aux1 = nettoBohr / labOffset;
                double floor = Math.Floor(aux1);
                double nbrRowsTot = floor + 1;
                double aux2 = nbrRowsTot / 2;
                double nbrRows = Math.Ceiling(aux2);
                double nbrRowsOffset = Math.Floor(aux2);

                double nbrHolesRad = 360 / deg;
                double aNull = deg - deg;

                double RPD = 60 / lab;
                double Netto = ((nettoBohr / lab) + 1) / RPD;
                double LastL = (Netto - Math.Floor(Netto)) * RPD;
                double LetzteL = ((Netto * 15) * lab);

                #endregion

                #region -- Warnungen --
                if (lab == 0)
                {
                    MessageBox.Show("Reihen-Abstand darf nicht 0 sein", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                if (deg == 0)
                {
                    MessageBox.Show("Lochabstand zu klein", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (mat == 0)
                {
                    MessageBox.Show("Materialstärke darf nicht 0 sein", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (anz == 0)
                {
                    MessageBox.Show("Anzahl darf nicht 0 sein", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (secur == 0)
                {
                    MessageBox.Show("Sicherheitshöhe darf nicht 0 sein", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (durch == 0)
                {
                    var Nachricht = MessageBox.Show("Duchbruchtiefe bei 0", "Achtung!", MessageBoxButton.OKCancel, MessageBoxImage.Warning);

                    if (Nachricht == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                }

                if (spinRPM == 0)
                {
                    var Nachricht = MessageBox.Show("Drehzahl bei 0", "Achtung!", MessageBoxButton.OKCancel, MessageBoxImage.Warning);

                    if (Nachricht == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                }

                if (drillDia == 0)
                {
                    MessageBox.Show("Kein Bohrdurchmesser ausgewählt", "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                #endregion

                await Task.Run(() =>
                {
                    #region -- Logic --

                    using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                    {
                        writer.WriteLine("(" + fileName + ")");
                        writer.WriteLine("(dd=" + diaInt + " | h=" + height + " | netto=" + nettoBohr + " | t=" + mat + " | °rad=" + deg + " | #rad=" + nbrHolesRad + " | x=" + lab + " | #rows=" + nbrRows + "/" + nbrRowsOffset + " | spin=" + spinRPM + ")");
                        writer.WriteLine("(Verschiebung xNull: " + xVersch + ")");
                        writer.WriteLine("(Anzahl Formen: " + anz + " " + "Abstand: " + xAbst + ")");
                        writer.WriteLine("(T1 D=1. CR=0. TAPER=140DEG-Bohrer)");
                        writer.WriteLine("G90 G94 G91.1 G40 G49 G17");
                        writer.WriteLine("G21");
                        writer.WriteLine("G28 G91 Z0.");
                        writer.WriteLine("G90");                         //
                                                                         //writer.WriteLine("M5");                         '(=Spindelstopp, braucht es nicht)
                        writer.WriteLine("T1 M6");                       //Werkzeug 1, M6=Wz-Wechsel
                        writer.WriteLine("S" + spinRPM + " M3");         //Spindeldrehzahl
                        writer.WriteLine("G54");                          //Werkstücknullpunkt"

                        Spind1 = true;
                        Spind2 = true;
                        Spind3 = true;
                        Spind4 = true;
                        aPos = aNull;
                        Currentx = 0;

                        writer.WriteLine("m110"); //alle hoch
                        writer.WriteLine("m101"); //alle runter (one by one)
                        writer.WriteLine("m102");
                        writer.WriteLine("m103");
                        writer.WriteLine("m104");

                        writer.WriteLine("G0 A0");
                        writer.WriteLine("G0 X0");
                        writer.WriteLine("G43 Z" + (radExt + secur) + " H1");
                        writer.WriteLine("G98 G81 X" + xErste + " Z" + ((radExt + secur) - (mat + durch)) + " R" + (radExt + secur) + " F8000");

                        if (Netto <= 4)
                        {
                            if (Netto == Math.Floor(Netto)) //wenn gerade zahl (Note: VB Int() equivalent is Math.Floor())
                            {
                                if (Netto == 1)       //richtige spindeln auswählen
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind2 = false;
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Netto == 2)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Netto == 3)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                                // Note: Boolean.ToString() in C# outputs "True"/"False" (capitalized), matching VB.NET behavior.
                                writer.WriteLine("S1:" + Spind1 + " S2:" + Spind2 + " S3:" + Spind3 + " S4:" + Spind4 + "");

                                for (h = 1; h <= RPD; h++) //einmal bis 15, also ganzer durchgang (VB: For h = 1 To RPD)
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + xPos + "");
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                            }
                            else
                            {
                                writer.WriteLine("(Not Whole)"); //wenn ungerade
                                if (Math.Floor(Netto) == 0) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind2 = false;
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 1) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 2) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                                writer.WriteLine("(S1:" + Spind1 + " S2:" + Spind2 + " S3:" + Spind3 + " S4:" + Spind4 + ")");
                                for (h = 1; h <= (int)Math.Floor(LastL); h++) //zuerst mit allen möglichen bis zum letzten punkt..... (VB: For h = 1 To Int(LastL))
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("(LastL-Reihen)");
                                    writer.WriteLine("X" + xPos + "");
                                    Currentx = xPos;
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                                if (Math.Floor(Netto) == 1) //die richtigen spindeln hochnehmen (VB: Int(Netto))
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = false;
                                }
                                else if (Math.Floor(Netto) == 2) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = false;
                                }
                                else if (Math.Floor(Netto) == 3) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                                writer.WriteLine("(Rest nach spindelrückzug)");
                                xPos = Currentx;
                                // VB: For h = Int(LastL) To RPD - 1
                                // C#: Loop runs from floor(LastL) up to and including RPD - 1
                                for (h = (int)Math.Floor(LastL); h <= RPD - 1; h++) //den rest weiter bohren mit einer spindel weniger
                                {
                                    xPos = xPos + lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + (h + 1) + ")"); // Note: Original VB comment implies h starts from Int(LastL) and goes up to RPD-1. The row number is h+1.
                                    writer.WriteLine("X" + xPos + "");
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                            }
                        }
                        else if (Netto <= 8) // falls netto über 4 ist oder form über 236 lang
                        {
                            writer.WriteLine("(grösser 4)");
                            if (Netto == Math.Floor(Netto)) //gerade netto zahlen (VB: Int(Netto))
                            {
                                writer.WriteLine("(Gerade)");
                                for (h = 1; h <= RPD; h++) //einmal ganz durch fahren da es ja einer form über 4 ist (VB: For h = 1 To RPD)
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + xPos + "");
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                                if (Netto == 5)      //bohrer auswählen
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                }
                                else if (Netto == 6)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                }
                                else if (Netto == 7)
                                {
                                    writer.WriteLine("m104");
                                }
                                else if (Netto == 8)
                                {
                                    writer.WriteLine("(No Retraction!)");
                                }

                                for (h = 1; h <= RPD; h++) //mit ausgewählten bohrern ganz bohren (VB: For h = 1 To RPD)
                                {
                                    xPos = 240 + xErste + (h - 1) * lab; // Note the added 240 offset here
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + xPos + "");
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                            }
                            else //falls netto nicht gerade ist
                            {
                                writer.WriteLine("(ungerade)");
                                for (h = 1; h <= RPD; h++) //einmal ganz durch wie vorher (VB: For h = 1 To RPD)
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + xPos + "");
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }

                                writer.WriteLine("(Not Whole)"); //bohrer...
                                if (Math.Floor(Netto) == 4) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind2 = false;
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 5) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 6) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                                writer.WriteLine("(S1:" + Spind1 + " S2:" + Spind2 + " S3:" + Spind3 + " S4:" + Spind4 + ")");
                                // VB: For h = 1 To Int(LastL) + 1
                                // C#: Loop runs from 1 up to and including floor(LastL) + 1
                                for (h = 1; h <= (int)Math.Floor(LastL); h++) //bis zur letzten bohren mit allen möglichen
                                {
                                    xPos = 240 + xErste + (h - 1) * lab; // Note the added 240 offset here
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + xPos + "");
                                    Currentx = xPos;
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                                if (Math.Floor(Netto) == 4) //Bohrer (VB: Int(Netto))
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = false;
                                }
                                else if (Math.Floor(Netto) == 5) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = false;
                                }
                                else if (Math.Floor(Netto) == 6) // VB: Int(Netto)
                                {
                                    if (Spind4 == false)
                                    {
                                        writer.WriteLine("m104");
                                        Spind4 = false; // Corrected typo from original VB comment analysis (assuming it meant Spind4 = False)
                                    }
                                }
                                writer.WriteLine("(Rest nach spindelrückzug)");
                                xPos = Currentx;
                                // VB: For h = Int(LastL) To RPD - 2
                                // C#: Loop runs from floor(LastL) up to and including RPD - 2
                                for (h = (int)Math.Floor(LastL); h <= RPD - 2; h++) //den rest auffüllen
                                {
                                    xPos = xPos + lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + (h + 1) + ")"); // Note: Row number is h+1
                                    writer.WriteLine("X" + xPos + "");
                                    for (i = 1; i <= nbrHolesRad; i++) // VB: For i = 1 To nbrHolesRad
                                    {
                                        aPos = aPos + deg;
                                        writer.WriteLine("A" + aPos + "");
                                    }
                                }
                            }
                        }
                        writer.WriteLine("G80");
                        writer.WriteLine("m110"); //alle spindeln hoch
                        writer.WriteLine("G28 G91 Z0.");    //
                        writer.WriteLine("G90");
                        writer.WriteLine("A0.");
                        writer.WriteLine("G28 G91 X0.");
                        //writer.WriteLine("G90");
                        writer.WriteLine("M5");
                        writer.WriteLine("M30");

                        writer.Close();
                    }
                    #endregion
                });
                #region -- speichern --

                Properties.Settings.Default.pathInput = directory;

                if (IsChecked == true)
                {
                    Properties.Settings.Default.height = double.Parse(_height.Text);
                    Properties.Settings.Default.lab = double.Parse(_lab.Text);
                    Properties.Settings.Default.mat = double.Parse(_mat.Text);
                    Properties.Settings.Default.xErste = double.Parse(_xErste.Text);
                    Properties.Settings.Default.rabo = double.Parse(_rabo.Text);
                    Properties.Settings.Default.deg = double.Parse(_deg.Text);
                    Properties.Settings.Default.durch = double.Parse(_durch.Text);
                    Properties.Settings.Default.spinRPM = _spinRPM.Value;
                    Properties.Settings.Default.xVersch = double.Parse(_xVersch.Text);
                    Properties.Settings.Default.diaInt = double.Parse(_diaInt.Text);
                    Properties.Settings.Default.secur = double.Parse(_secur.Text);
                    Properties.Settings.Default.xAbst = double.Parse(_xAbst.Text);

                    Properties.Settings.Default.IsChecked = true;

                    string selectedMode = OptionA.IsChecked == true ? "0.8" :
                                OptionB.IsChecked == true ? "1.0" :
                                OptionC.IsChecked == true ? "1.3" :
                                OptionD.IsChecked == true ? "1.5" :
                                OptionE.IsChecked == true ? "2.0" :
                                OptionF.IsChecked == true ? "3.0" :
                                "None";

                    Properties.Settings.Default.drillDia = selectedMode;

                    Properties.Settings.Default.Save();
                }
                else
                {
                    Properties.Settings.Default.IsChecked = false;
                    ClearInputs(null, null);
                }
                #endregion

                MessageBox.Show("Code Erstellt unter: " + filePath, "Erfolgreiche Eingabe", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            #endregion
        }

    }

}
