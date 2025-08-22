#Vérifie la présence de "MakeCAB.exe" sous "C:\Windows\System32\"
If(!(Test-Path -Path "$Env:SystemRoot\System32\makecab.exe")){[void][System.Windows.Forms.MessageBox]::Show("MakeCAB ne semble pas présent sur l'ordinateur.`nAssurez-vous de la présence de l'exécutable :`n`n""C:\Windows\System32\makecab.exe"""," Erreur","0","16"); Exit}

####################################
##  Masque la console Powershell  ##
####################################

Try{
    Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
 
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)  | Out-Null
}
Catch {
    Throw "Failed to load Windows dll."
}


#############################################
##  CHARGEMENT DES ASSEMBLY NET FRAMEWORK  ##
#############################################

$code = @"
using System;
using System.Drawing;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}

namespace ProgressBar
{
  [ToolboxBitmap(typeof(System.Windows.Forms.ProgressBar))]
  public class VistaProgressBar : System.Windows.Forms.ProgressBar
  {
    public delegate void StateChangedHandler(object source, vState State);
    
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern uint SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);
    
    private vState _State = vState.Normal;
    
    public enum vState { Normal, Pause, Error }
    
    private const int WM_USER = 0x400;
    private const int PBM_SETSTATE = WM_USER + 16;
    
    private const int PBST_NORMAL = 0x0001;
    private const int PBST_ERROR = 0x0002;
    private const int PBST_PAUSED = 0x0003;
    
    [Category("Behavior")]
    [Description("Event raised when the state of the Control is changed.")]
    public event StateChangedHandler StateChanged;
    
    [Category("Behavior")]
    [Description("This property allows the user to set the state of the ProgressBar.")]
    [DefaultValue(vState.Normal)]
    public vState State
    {
      get
      {
        if (Environment.OSVersion.Version.Major < 6)
          return vState.Normal;
        if (this.Style == System.Windows.Forms.ProgressBarStyle.Blocks) return _State;
        else return vState.Normal;
      }
      set
      {
        _State = value;
        if (this.Style == System.Windows.Forms.ProgressBarStyle.Blocks)
          ChangeState(_State);
      }
    }
    private void ChangeState(vState State)
    {
      if (Environment.OSVersion.Version.Major > 5) {
        SendMessage(this.Handle, PBM_SETSTATE, PBST_NORMAL, 0);

        switch (State) {
          case vState.Pause:
            SendMessage(this.Handle, PBM_SETSTATE, PBST_PAUSED, 0);
            break;
          case vState.Error:
            SendMessage(this.Handle, PBM_SETSTATE, PBST_ERROR, 0);
            break;
          default:
            SendMessage(this.Handle, PBM_SETSTATE, PBST_NORMAL, 0);
            break;
        }
        if (StateChanged != null)
          StateChanged(this, State);
      }
    }
    protected override void WndProc(ref System.Windows.Forms.Message m)
    {
      if (m.Msg == 15)
        ChangeState(_State);
      base.WndProc(ref m);
    }
  }
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies System.Windows.Forms, System.Drawing

#############################################
##  DECLARATION DES FONCTIONS & VARIABLES  ##
#############################################

#Déclaration des répertoires courant et de travail
$CurrentScriptPath = $myInvocation.MyCommand.Definition
$CurrentDir = [System.IO.Path]::GetDirectoryName($CurrentScriptPath)
$WorkingDir = "$Env:SystemDrive\TEMP\src\"

#Déclaration de la fonction de suppression des accents
#Function FormatText{Param ([String]$String)[Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))}

#Déclaration de la fonction de réinitialisation des barres de progression et de statut
Function BarsRefresh{
  $progressB.BringToFront()
  $progress.SendToBack()
  Start-Sleep -Milliseconds 50
  $progress.Value = 0
  $progress.State = 'Normal'
  Start-Sleep -Milliseconds 50
  $progress.BringToFront()
  $progressB.SendToBack()
  If($($listBox.Items.Count -gt 0)){$statusBar.Text = (" Liste de $($listBox.Items.Count) objets")}
  Else{$statusBar.Text = " Prêt"}
  If($($listBox.Items.Count -gt 0)){$button_ok.Enabled = $True}
  Else{$button_ok.Enabled = $False}
}

#Ajout d'un fichier/Dossier au TreeView
Function Add-Node{ 
  Param ($selectedNode,$name)
  $newNode = New-Object System.Windows.Forms.TreeNode
  $newNode.Name = $name
  $newNode.Text = $name
  $selectedNode.Nodes.Add($newNode) | Out-Null
  return $newNode
}

#Gestion des niveaux d'arbrescence du TreeView
Function Get-NextLevel{
  Param ($selectedNode,$Sub)
  If($Sub.Name -eq $Null){$SName = (Split-Path -Path $Sub -leaf)}
  Else{$SName = $Sub.Name}
  $array = @(Get-Item -Path "$Sub\*" -Force)
  If((Get-ChildItem -Path $Sub -Force) -eq $Null){$node = Add-Node $selectedNode $SName}
  Else{
    $node = Add-Node $selectedNode $SName
    $array | ForEach-Object {Get-NextLevel $node $_}
  }
}

#Gestion des niveaux d'arbrescence dans le fichier ddf
Function ddfTree{
  Param ($SubDDFTree)
  ForEach ($item in $SubDDFTree){
    (".Set DestinationDir=$('"' + $item.FullName + '"')") -replace ([regex]::Escape($WorkingDir),"") | Add-Content -path $ddf -Encoding Default
    (Get-ChildItem -LiteralPath $($item.FullName) -Force -File) | ForEach-Object{'"' + $_.FullName + '"' -replace ([regex]::Escape($WorkingDir),'')} | Add-Content -path $ddf -Encoding Default
    $SubDDFTree=@(Get-ChildItem -LiteralPath $($item.FullName) -Force -Directory)
    ddfTree $SubDDFTree
  }
}

#Créer des fichiers ".placeholder" dans les répertoires vides pour forcer MakeCAB à les inclure à l'archive Cabinet
Function PreserveEmptyDirectory {
    param(
        [string]$RootPath
    )
    Get-ChildItem -Path $RootPath -Recurse -Directory | ForEach-Object {
        if (-not (Get-ChildItem -Path $_.FullName -Force)) {
            New-Item -Path $_.FullName -Name ".placeholder" -ItemType File -Force | Out-Null
        }
    }
}

#Format automatique des octets
Function DisplayInBytes($num){
  $suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
  $index = 0
  while ($num -gt 1kb){
    $num = $num / 1kb
    $index++
  } 
  "{0:N1} {1}" -f $num, $suffix[$index]
}

#Création du fichier de configuration masqué
$SS64 = @"
;*** Cabinet Maker – MakeCAB Directive

.Set FailOnMissingSource=On
.Set Cabinet=On
"@
$ddf = "$CurrentDir\SS64.ddf"
New-Item -Path $ddf –Type File -Force | Out-Null
Set-ItemProperty -Path $ddf -Name Attributes -Value ([System.IO.FileAttributes]::Hidden) -Force | Out-Null
Set-Content -Path $ddf -Value $SS64 -Encoding Default

#Activation des effets visuels
[Windows.Forms.Application]::EnableVisualStyles()


###########################################
##  CONFIGURATION DU FORMULAIRE WINDOWS  ##
###########################################

#Creation de la form principale
$form = New-Object System.Windows.Forms.Form
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $False
$form.MinimizeBox = $True
$form.Size = New-Object System.Drawing.Size(1459,775)
$form.Text = "Cabinet Maker"
#$form.Topmost = $True
$form.Icon = [System.IconExtractor]::Extract("cabview.dll", 0, $true)
$form.ShowInTaskbar = $True
#$form.FlatStyle ="Standard"
 
 
###########################
## AJOUT DES COMPOSANTS  ##
###########################

#Label Sélection
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = '20,55'
$label1.AutoSize = $True
$label1.Text = "Glisser-Déposer ou Ajouter les dossiers et fichiers à archiver :"
$label1.Font = New-Object System.Drawing.Font("arial",10,[System.Drawing.FontStyle]::Bold);

#Bouton Ajouter dossiers
$button_addFolder = New-Object System.Windows.Forms.Button
$button_addFolder.Location = '1101,52'
$button_addFolder.Size = '105,25'
$button_addFolder.text = "Ajouter dossier"

#Bouton Ajouter fichiers
$button_addFile = New-Object System.Windows.Forms.Button
$button_addFile.Location = '1210,52'
$button_addFile.Size = '105,25'
$button_addFile.text = "Ajouter fichiers"

#Bouton Effacer la liste
$button_supp = New-Object System.Windows.Forms.Button
$button_supp.Location = '1319,52'
$button_supp.Size = '105,25'
$button_supp.Text = "Effacer la liste"

#ListBox Sélection
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = '20,80'
$listBox.Height = 350
$listBox.Width = 1403
$listBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$listBox.IntegralHeight = $False
$listBox.AllowDrop = $True

#GroupBox Compression
$GroupBox1 = New-Object System.Windows.Forms.GroupBox
$GroupBox1.Location = '20,460'
$GroupBox1.size = '300,110'
$GroupBox1.text = " Compression : "
$GroupBox1.Font = New-Object System.Drawing.Font("arial",10,[System.Drawing.FontStyle]::Bold)
    
#Radio Activer
$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Location = '20,35'
$RadioButton1.size = '110,20'
$RadioButton1.Checked = $True 
$RadioButton1.Text = "Activer"
$RadioButton1.Font = New-Object System.Drawing.Font("arial",10)

#Radio Désactiver
$RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioButton2.Location = '20,65'
$RadioButton2.size = '110,20'
$RadioButton2.Checked = $False
$RadioButton2.Text = "Désactiver"
$RadioButton2.Font = New-Object System.Drawing.Font("arial",10)

#Label Type de compression
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = '150,35'
$label2.AutoSize = $True
$label2.Text = "Type :"
$label2.Font = New-Object System.Drawing.Font("arial",10)

#Combobox Type de compression
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = '210,33'
$comboBox1.Size = '67,19'
$comboBox1.Enabled = $True
$comboBox1.DropDownStyle = "DropDownList"
$comboBox1.Font = New-Object System.Drawing.Font("arial",10)

#Label Niveau de compression
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = '150,67'
$label3.AutoSize = $True
$label3.Text = "Niveau :"
$label3.Font = New-Object System.Drawing.Font("arial",10)

#Combobox Niveau de compression
$comboBox2 = New-Object System.Windows.Forms.ComboBox
$comboBox2.Location = '224,64'
$comboBox2.Size = '53,19'
$comboBox2.Enabled = $True
$comboBox2.DropDownStyle = "DropDownList"
$comboBox2.Font = New-Object System.Drawing.Font("arial",10)

#GroupBox Découper
$GroupBox2 = New-Object System.Windows.Forms.GroupBox
$GroupBox2.Location = '335,460'
$GroupBox2.size = '211,110'
$GroupBox2.text = " Découper en volumes : "
$GroupBox2.Font = New-Object System.Drawing.Font("arial",10,[System.Drawing.FontStyle]::Bold)
[Long]$Size=0

#Radio Non
$RadioButton3 = New-Object System.Windows.Forms.RadioButton
$RadioButton3.Location = '25,35'
$RadioButton3.size = '75,20'
$RadioButton3.Checked = $True 
$RadioButton3.Text = "Non"
$RadioButton3.Font = New-Object System.Drawing.Font("arial",10)

#Radio Oui
$RadioButton4 = New-Object System.Windows.Forms.RadioButton
$RadioButton4.Location = '110,35'
$RadioButton4.size = '75,20'
$RadioButton4.Checked = $False
$RadioButton4.Text = "Oui"
$RadioButton4.Font = New-Object System.Drawing.Font("arial",10)

#Combobox Découper en volumes
$comboBox3 = New-Object System.Windows.Forms.ComboBox
$comboBox3.Location = '20,65'
$comboBox3.Size = '170,20'
$comboBox3.Enabled = $False
$comboBox3.Text = ''
$comboBox3.Font = New-Object System.Drawing.Font("arial",10);
$comboBox3.Text = '2 Gio (Max)'

#GroupBox Répertoire de destination
$GroupBox4 = New-Object System.Windows.Forms.GroupBox
$GroupBox4.Location = '562,460'
$GroupBox4.size = '860,110'
$GroupBox4.text = " Répertoire de destination : "
$GroupBox4.Font = New-Object System.Drawing.Font("arial",10,[System.Drawing.FontStyle]::Bold)

#TextBox Répertoire de destination
$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = '20,50'
$textBox1.size = '798,20'
$textBox1.Enabled = $False
$textBox1.BorderStyle ="FixedSingle"
$textBox1.ForeColor = "Blue"
$textBox1.BackColor = "White"
$textBox1.Text = "$CurrentDir\archive.cab"
$textBox1.Font = New-Object System.Drawing.Font("arial",10,[System.Drawing.FontStyle]::Regular)
$textBox1.AutoCompleteMode = "Append"
$textBox1.AutoCompleteSource = "FileSystem"

#Bouton Parcourir...
$button_elli = New-Object System.Windows.Forms.Button
$button_elli.Location = '815,49'
$button_elli.Size = '25,25'
$button_elli.text = "..."
$button_elli.Font = New-Object System.Drawing.Font("arial",8);

#Barre de progression
$progress = New-Object ProgressBar.VistaProgressBar
$progress.Location = '20,610'
$progress.Size = '1403,23'
$progress.Name = 'progressBar'
$progress.Minimum = 0
$progress.Maximum=100
$progress.Value = 0
$progress.BringToFront()
$progressB = New-Object System.Windows.Forms.ProgressBar
$progressB.Location = '20,610'
$progressB.Size = '1403,23'
$progressB.Name = 'progressBarB'
$progressB.Value = 0
$progressB.SendToBack()

#Bouton OK
$button_ok = New-Object System.Windows.Forms.Button
$button_ok.Text = "OK"
$button_ok.Enabled = $False
$button_ok.Location = '1088,670'
$button_ok.Size = '160,35'
#$button_ok.DialogResult=[System.Windows.Forms.DialogResult]::OK

#Bouton Quitter
$button_quit = New-Object System.Windows.Forms.Button
$button_quit.Text = "Fermer"
$button_quit.Location = '1264,670'
$button_quit.Size = '160,35'
#$button_quit.DialogResult=[System.Windows.Forms.DialogResult]::Cancel

#Barre de Statut
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Text = " Prêt"

#Définition du formulaire
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
$formTree = New-Object System.Windows.Forms.Form 
$formTree.Text = " Ajouter l'arborecence ?" 
$formTree.Name = "Arborecence" 
$formTree.DataBindings.DefaultDataSourceUpdateMode = 0
$formTree.ClientSize = New-Object System.Drawing.Size(400,600)
$formTree.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$formTree.Icon = [System.IconExtractor]::Extract("cabview.dll", 0, $true)
$formTree.StartPosition = "CenterScreen"
$formTree.MaximizeBox = $False
$formTree.MinimizeBox = $False
$formTree.Topmost = $True

#Définition du TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(370,380)
$treeView.Name = "treeView" 
$treeView.Location = New-Object System.Drawing.Size(15,31)
$treeView.DataBindings.DefaultDataSourceUpdateMode = 0
$treeView.HideSelection = $False
$treeView.TabIndex = 0

#Suppression les bordures
#$treeView.Dock = [System.Windows.Forms.DockStyle]::Fill

#Définition du titre de la textBox Elément sélectionné
$label = New-Object System.Windows.Forms.Label
$label.Name = "label" 
$label.Location = New-Object System.Drawing.Size(14,485)
$label.Size = New-Object System.Drawing.Size(150,20)
$label.Text = "Dossier/Fichier sélectionné :"

#Définition du titre de la textBox Nombre de fichiers
$labelNbFiles = New-Object System.Windows.Forms.Label
$labelNbFiles.Name = "labelNbFiles" 
$labelNbFiles.Location = New-Object System.Drawing.Size(14,430)
$labelNbFiles.Size = New-Object System.Drawing.Size(105,20)
$labelNbFiles.Text = "Nombre de fichiers :"

#Définition du titre de la textBox Nombre de dossiers
$labelNbDir = New-Object System.Windows.Forms.Label
$labelNbDir.Name = "labelNbDir" 
$labelNbDir.Location = New-Object System.Drawing.Size(220,430)
$labelNbDir.Size = New-Object System.Drawing.Size(130,20)
$labelNbDir.Text = "Nombre de dossiers :"

#Définitation de la textBox Sélection
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Name = "textbox" 
$textbox.Location = New-Object System.Drawing.Size(15,505)
$textbox.Size = New-Object System.Drawing.Size(370,20)
$textbox.Text = ""
$textbox.Enabled = $False

#Bouton Ajouter l'arborescence
$ButtonTree = New-Object System.Windows.Forms.Button
$ButtonTree.Location = "267,550"
$ButtonTree.Size = "118,33"
$ButtonTree.Text = "Arborescence"

#Bouton Ajouter les fichiers enfants
$ButtonSubFiles = New-Object System.Windows.Forms.Button
$ButtonSubFiles.Location = "141,550"
$ButtonSubFiles.Size = "118,33"
$ButtonSubFiles.Text = "Fichiers seulement"

#Bouton Annuler
$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Location = "15,550"
$ButtonCancel.Size = "118,33"
$ButtonCancel.Text = "Annuler"
[boolean]$script:ExitTreeForm = $True

#Enregistre l'état initial du formulaire
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState 
$InitialFormWindowState = $formTree.WindowState 

#Définitation de la textBox Nombre de fichiers
$textboxNbDir = New-Object System.Windows.Forms.TextBox
$textboxNbDir.Name = "textboxNbDir" 
$textboxNbDir.Location = New-Object System.Drawing.Size(122,428)
$textboxNbDir.Size = New-Object System.Drawing.Size(50,20)
$textboxNbDir.Text = 0
$textboxNbDir.Enabled = $False

#Définitation de la textBox Nombre de dossiers
$textboxNbFiles = New-Object System.Windows.Forms.TextBox
$textboxNbFiles.Name = "textboxNbFiles" 
$textboxNbFiles.Location = New-Object System.Drawing.Size(335,428)
$textboxNbFiles.Size = New-Object System.Drawing.Size(50,20)
$textboxNbFiles.Text = 0
$textboxNbFiles.Enabled = $False

#Bouton Développer/Réduire l'arborescence
$Button = New-Object System.Windows.Forms.Button
$Button.Location = "306,10"
$Button.Size = "80,22"
$Button.Text =　"Développer"
$open_close = {
  If($Button.Text -eq "Développer"){
    $TreeView.ExpandAll()
    $button.Text = "Réduire"
    $treeView.focus()
  }
  Else{
    $treeView.CollapseAll()
    $button.Text = "Développer"
    $textbox.Text = $treeNodes.Name
    $treeView.focus()
  }
}


###############################
##  CHARGEMENT DES COMBOBOX  ##
###############################

#Chargement de la comboBox Type de compression
[void]($comboBox1.Items.Add("LZX"))
[void]($comboBox1.Items.Add("MSZIP"))
$comboBox1.SelectedIndex=0

#Chargement de la comboBox Niveau de compression
$Levels = @(15,16,17,18,19,20,21)
ForEach($Lev in $Levels){
  [void]($comboBox2.Items.add($Lev))
  If($Lev -eq 21){Break}
  Else{$Lev = $Lev + 1}
}
$comboBox2.SelectedIndex=6

#Chargement de la comboBox Découper en volumes
$Cut = @('1 Mio','5 Mio','10 Mio','25 Mio','50 Mio','75 Mio','100 Mio','250 Mio','500 Mio','700 Mio','800 Mio','1 Gio','1,5 Gio','2 Gio (Max)')
ForEach($Multi in $Cut){[void]($comboBox3.Items.add($Multi))}


################################
##  INSERTION DES COMPOSANTS  ##
################################

# Ajout des composants a la Form
$form.SuspendLayout()
$form.Controls.Add($label1)
$form.Controls.Add($listBox)
$form.controls.add($button_addFolder)
$form.controls.add($button_addFile)
$form.Controls.Add($button_supp)
$form.Controls.AddRange($GroupBox1)
$GroupBox1.Controls.AddRange(@($Radiobutton1,$RadioButton2,$label2,$comboBox1,$label3,$comboBox2))
$form.Controls.AddRange($GroupBox2)
$GroupBox2.Controls.AddRange(@($comboBox3,$RadioButton3,$RadioButton4))
$form.Controls.AddRange($GroupBox4)
$GroupBox4.Controls.AddRange(@($TextBox1,$button_elli))
$form.Controls.Add($progressB)
$form.Controls.Add($progress)
$form.Controls.Add($button_ok)
$form.Controls.Add($button_quit)
$form.Controls.Add($statusBar)
$formTree.Controls.AddRange(@($OnLoadForm_StateCorrection,$treeView,$label,$textbox,$Button,$labelNbFiles,$textboxNbFiles,$labelNbDir,$textboxNbDir,$ButtonTree,$ButtonSubFiles,$ButtonCancel))
$form.ResumeLayout()


###################################
##  DESCRIPTION DES EVENEMENTS   ##
###################################

$button_ok_Click = {
  BarsRefresh
  #Initialise la variable $Exit à $False
  $Exit = $False
  #Aiguiller la commande à exécuter en fonction de la valeur de la comboBox Découper en volumes
  [String]$Size = $comboBox3.Text
  Switch -regex ($Size){
    '^1 Mio$'          {[Long]$Size=1048576; Break}
    '^5 Mio$'          {[Long]$Size=5242880; Break}
    '^10 Mio$'         {[Long]$Size=10485760; Break}
    '^25 Mio$'         {[Long]$Size=26214400; Break}
    '^50 Mio$'         {[Long]$Size=52428800; Break}
    '^75 Mio$'         {[Long]$Size=78643200; Break}
    '^100 Mio$'        {[Long]$Size=104857600; Break}
    '^250 Mio$'        {[Long]$Size=262144000; Break}
    '^500 Mio$'        {[Long]$Size=524288000; Break}
    '^700 Mio$'        {[Long]$Size=734003200; Break}
    '^800 Mio$'        {[Long]$Size=838860800; Break}
    '^1 Gio$'          {[Long]$Size=1073741824; Break}
    '^1,5 Gio$'        {[Long]$Size=1610612736; Break}
    '^2 Gio \(Max\)$'  {[Long]$Size=0; Break}
    #Une archive Cabinet ne peut stocker au maximum que 65 535 fichiers pour un volume total ne pouvant excéder 1,99 Gio. L'archive peut-être découpée en multiples volumes de tailles prédéterminées correspondant à un multiple de 512 : 1,99 Gio <> 2136746229,76 octets | 2136746229/512=4173332,48 => 4173332*512 = 2136745984 => 2136745472 => 2136263680
    #Pour $Size commençant par des chiffres, espacé ou non de l'unité octets (débutant ou non par une majuscule & singulier ou pluriel) => Si $Size est inférieur à 512 retourner message d'erreur, sinon formater $Size en tant que plus grand multiple de 512
    '^[0-9]*(( )|())((O|octets$)|(O|octet$))'  {[Long]$Size = $Size -replace '(( )|())(([O|o]ctets$)|([O|o]ctet$))',""; If($Size -lt 512){[System.Windows.Forms.MessageBox]::Show("Saisir une valeur suivie de l'unité octets, Mio ou Gio (< 2 Gio)"," Erreur","0","16")}Else{$Size = $Size - ($Size%512)}; Break}
    #Pour $Size commençant par des chiffres, espacé ou non de l'unité Mio (débutant ou non par une majuscule) => Mutiplication de la valeur par 1048576 (<> 1 Mio) pour obtenir la valeur en Mégaoctets
    '^[0-9]*(( )|())[M|m]io$'                  {[Long]$Size = $Size -replace '(( )|())[M|m]io$',""; $Size = $Size * 1048576; Break}
    #Pour $Size commençant par des chiffres, espacé ou non de l'unité Gio (débutant ou non par une majuscule) => Mutiplication de la valeur par 1073741824 (<> 1 Gio) pour obtenir la valeur en Gigaoctets
    '^[0-9]*(( )|())[G|g]io$'                  {[Long]$Size = $Size -replace '(( )|())[G|g]io$',""; $Size = $Size * 1073741824; If($Size -gt 2126008320){$Size = 0}; Break}
    #Pour $Size commençant par des chiffres, espacé ou non de l'unité Gio (débutant ou non par une majuscule) => Mutiplication de la valeur par 1099511531398 (<> 1 Tio) pour obtenir la valeur en Gigaoctets
    #'^[0-9]*(( )|())[T|t]io$'                  {[Long]$Size = $Size -replace '(( )|())[T|t]io$',""; $Size = $Size * 1099511531398; Break}
    #Pour tout autre résultat, afficher un message d'erreur sur le format et passer la variable $Exit à $True
    Default                                    {[System.Windows.Forms.MessageBox]::Show("Saisir une valeur suivie de l'unité octets, Mio ou Gio (< 2 Gio)"," Erreur","0","16"); $Exit = $True;Break}
  }
  #Définit le chemin de destination du fichier CAB en fonction de la textBox Répertoire de destination
  $DestRep = $textBox1.Text.Substring(0, $textBox1.Text.LastIndexOf('\'))
  #Définit le nom du fichier CAB en fonction de la textBox Répertoire de destination
  $CABName = ($textBox1.Text.Substring($textBox1.Text.LastIndexOf('\'), ($textBox1.Text.Length - $textBox1.Text.LastIndexOf('\')))) -replace '\\',''
  #Si le total des objets listés dans la ListeBox est supérieur à la taille maximum de l'archive (contrainte du format ou découpage souhaité), on ajoute une étoile avant l'extension de l'archive de sortie
  $RealSize = (Get-ChildItem $listBox.Items -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
  If(($RealSize -gt $Size) -Or ($RealSize -gt 2126008320)){
    [string]$ShortName = [System.IO.Path]::GetFileNameWithoutExtension($CABName)
    If($RadioButton4.Checked){$CABName = $CABName.Replace($ShortName,$($ShortName + '*'))}
  }
  #Vérification de l'existance d'une archive Cabinet du même nom dans le répertoire de destination
  If(Test-Path -Path (($DestRep + $CABName) -replace '\*','1')){
    $OverWrite = [System.Windows.Forms.MessageBox]::Show("Le fichier $CABName existe déjà dans le répertoire de destination.`nVoulez-vous l'écraser ?"," Ecraser ?","1","48")
    If($OverWrite -ne 'OK'){
      $Exit = $True
      #Mise à jour de la bar de progression et de la bar de statut
      $progress.State = 'Error'
      $statusBar.Text = " Création de l'archive interrompue par l'utilisateur"
      [System.Windows.Forms.Application]::DoEvents()
      Start-Sleep -Milliseconds 50
      $progress.Value = 100
    }
  }
  #Si la variable $Exit n'est pas égale à $True, alors
  If($Exit -eq !$True){
    $progress.State = 'Normal'
    $statusBar.Text = " Création de l'archive en cours..."
    $progress.Value = 15
    #Si la radio Activer Compression est sélectionnée, ...
    If($RadioButton1.Checked -eq $true){
      #alors la variable $EnabComp = "On"
      $EnabComp = "On"
      #Récupération de la valeur de la comboBox Type de compression dans la variable $Type
      $Type = $comboBox1.Text
      #Récupération de la valeur de la comboBox Niveau de compression dans la variable $Level
      If($Type -eq "LZX"){
        $Level = $comboBox2.Text
        #Stockage sous forme de tableau des données de configuration du fichier SS64.ddf
        $Sets = @(".Set Compress=$EnabComp",".Set MaxDiskSize=$Size",".Set CompressionType=$Type",".Set CompressionMemory=$Level",".Set DiskDirectoryTemplate=$DestRep",".Set CabinetNameTemplate=$CABName",".Set SourceDir=$WorkingDir",'.Set MaxErrors=1','.Set FolderFileCountThreshold=65535','.Set CabinetFileCountThreshold=65535','','',';*** Liste des fichiers à compresser dans le l''archive CAB','')
      }
      Else{$Sets = @(".Set Compress=$EnabComp",".Set MaxDiskSize=$Size",".Set CompressionType=$Type",".Set DiskDirectoryTemplate=$DestRep",".Set CabinetNameTemplate=$CABName",".Set SourceDir=$WorkingDir",'.Set MaxErrors=1','.Set FolderFileCountThreshold=65535','.Set CabinetFileCountThreshold=65535','','',';*** Liste des fichiers à compresser dans le l''archive CAB','')}
    }
    #Sinon, si la radio Désactiver Compression est sélectionnée, alors la variable $EnabComp = "Off"
    ElseIf($RadioButton2.Checked -eq $true){
    $EnabComp = "Off"
    #Stockage sous forme de tableau des données de configuration du fichier SS64.ddf
    $Sets = @(".Set Compress=$EnabComp",".Set MaxDiskSize=$Size",".Set DiskDirectoryTemplate=$DestRep",".Set CabinetNameTemplate=$CABName",".Set SourceDir=$WorkingDir",'.Set MaxErrors=1','.Set FolderFileCountThreshold=65535','.Set CabinetFileCountThreshold=65535','','',';*** Liste des fichiers à compresser dans le l''archive CAB','')
  }
  #Insertion des données de configuration au fichier SS64.ddf
  $Sets | ForEach-Object{$_} | Add-Content -Path $ddf -Encoding Default
  #Création du dossier sources temporaire où copier les données à archiver
  [void](New-Item $WorkingDir -ItemType Directory -Force -ErrorAction SilentlyContinue)
  #Masque le dossier sources temporaire
  #Set-ItemProperty -Path $WorkingDir -Name Attributes -Value ([System.IO.FileAttributes]::Hidden) -Force | Out-Null
  #Copie le fichiers sélectionnés via la listBox Sélection vers le dossier sources temporaire en supprimant les caractères spéciaux des noms de fichiers
  ForEach ($item in $listBox.Items){
    #Copie les fichiers présents dans la listbox sans les caractères spéciaux
    $File = ([string]$File = ($item.Substring($item.LastIndexOf('\'), ($item.Length - $item.LastIndexOf('\')))) -replace '\\','')
    If($File -ne '*'){
      Copy-Item -Path $item -Destination $($WorkingDir + $File) -Force}
      #Copie les arborscneces présentes dans la listbox, puis renomme les fichiers s'ils contiennent des caractères spéciaux
      Else{
        Copy-Item -Path $($item.Substring(0,$item.Length-2)) -Recurse -Force -Destination $WorkingDir
        Get-ChildItem -path $item -recurse | Foreach-Object {If(([string]$_.name) -ne $_.name){Rename-Item -Path $_.fullname -newname ([string]$_.name)}}
    }
  }
  PreserveEmptyDirectory -RootPath $WorkingDir
  $progress.Value = 50
  #Insertion des noms de fichiers contenus dans le dossier sources temporaire au fichier SS64.ddf
  ".Set DestinationDir=" | Add-Content -path $ddf -Encoding Default
  Get-ChildItem -Path $WorkingDir -File -Name -Force | ForEach-Object{'"' + $_ + '"'} | Add-Content -path $ddf -Encoding Default
  $array=@(Get-ChildItem -LiteralPath $WorkingDir -Force -Directory)
  ddfTree $array
  #Création de l'archive Cabinet à l'aide du fichier de configuration SS64.ddf
  $output = makecab.exe /F "$CurrentDir\SS64.ddf" 2>&1
  If ($LASTEXITCODE -ne 0) {
      $progress.State = 'Error'
      $statusBar.Text = " Erreur lors de la création de l'archive"
      [System.Windows.Forms.Application]::DoEvents()
      Start-Sleep -Milliseconds 50
      $progress.Value = 100
      $logPath = "$CurrentDir\CabinetMaker_Error_$(Get-Date -Format 'yyyy-MM-dd_HH.mm.ss').log"
      $logContent = "Erreur survenue le $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'):`n`n"
      $logContent += ($output | Out-String) # Out-String garantit un formatage correct
      Set-Content -Path $logPath -Value $logContent -Encoding UTF8
      $summary = $output | Select-Object -Last 10
      $summaryMessage = "Une erreur est survenue.`n`n"
      $summaryMessage += "Résumé de l'erreur :`n"
      $summaryMessage += "--------------------`n"
      $summaryMessage += ($summary | Out-String)
      $summaryMessage += "`nLes détails complets ont été enregistrés dans le fichier :`n$logPath"
      [System.Windows.Forms.MessageBox]::Show($summaryMessage, "Erreur retournée par makecab.exe", "0", "16")
  }
  #Sinon...
  Else{
    $progress.Value = 75
    #Suppression des fichiers de rapport makecab et du dossier de sources temporaire
    [void](Remove-Item "$CurrentDir\setup.rpt" -Force -Recurse -ErrorAction SilentlyContinue)
    [void](Remove-Item "$CurrentDir\setup.inf" -Force -Recurse -ErrorAction SilentlyContinue)
    [void](Remove-Item $WorkingDir -Force -Recurse -ErrorAction SilentlyContinue)
    #Suppression de la liste des items du listBox Sélection, mise à jour de la bar de progression et de la bar de statut
    $progress.Value = 100
    $listBox.Items.Clear()
    $statusBar.Text = " Archive terminée"
    #Désactivation du bouton OK
    $button_ok.Enabled = $False
  }
  #Quoiqu'il arrive, on efface le contenu du fichier de configuration SS64.ddf et on le réinitialise
  Clear-Content "$CurrentDir\SS64.ddf"
  Set-Content -Path $ddf -Value $SS64 -Encoding Default
  }
}


$button_quit_Click = {$form.Close()}


$button_supp_Click = {
  $listBox.Items.Clear()
  BarsRefresh
}


$button_addFile.add_click({
  #Réinitialisation la barre de progrossion et mise à jour de la bar de statut
  BarsRefresh
  #Création de l'objet
  $open = New-Object System.Windows.Forms.OpenFileDialog
  #Initialisation du chemin par défaut.
  $open.initialDirectory = $CurrentDir
  #Titre de la boite de dialogue
  $open.Title = "Sélectionner les fichiers à archiver"
  #Activation de la multi-sélection
  $open.Multiselect = $True
  #Récupére le chemin du raccourci et non celui de la cible
  $open.DereferenceLinks = $False
  $open.Filter = "All Files|*"
  #Affiche la fenêtre d'ouverture de fichier.
  $Sel = $open.ShowDialog()
  #Si le bouton "OK" est cliqué, alors...
  If($Sel -eq "OK"){#Pour chaque fichier sélectionné...
    ForEach($fullFilePath in $open.FileNames){
    #On vérifie qu'il n'est pas déjà présent dans la TextBox avant de l'afficher
    If(!($listBox.Items -Contains $fullFilePath)){
      $listBox.Items.Add($fullFilePath)
      $statusBar.Text = (" Liste de $($listBox.Items.Count) objets")}
      If($($listBox.Items.Count -gt 0)){$button_ok.Enabled = $True}
    }
  }
})


$button_addFolder.add_click({
  #Réinitialisation la barre de progrossion et mise à jour de la bar de statut
  BarsRefresh
  #Création de l'objet
  $open = New-Object System.Windows.Forms.FolderBrowserDialog
  $open.ShowNewFolderButton = $False
  #Titre de la boite de dialogue
  $open.Description = "Sélectionner les dossiers à archiver"
  #$open.Filter = "All Files|*"
  #Affiche la fenêtre d'ouverture de fichier.
  $Sel = $open.ShowDialog()
  #Si le bouton "OK" est cliqué, alors...
  If($Sel -eq "OK"){
    #Pour chaque fichier sélectionné...
    ForEach($fullFilePath in $open.SelectedPath){
      #On vérifie qu'il n'est pas déjà présent dans la TextBox avant de l'afficher
      If(!($listBox.Items -Contains "$fullFilePath\*")){
        $listBox.Items.Add($fullFilePath + "\*")
        $statusBar.Text = (" Liste de $($listBox.Items.Count) objets")
      }
      If($($listBox.Items.Count -gt 0)){$button_ok.Enabled = $True}
    }
  }
})


$listBox_DragOver = [System.Windows.Forms.DragEventHandler]{
  $script:ExitTreeForm = $True
  If($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)){$_.Effect = 'Copy'}
  Else{$_.Effect = 'None'}
}


$listBox_DragDrop = [System.Windows.Forms.DragEventHandler]{
  #Pour chaque élément copié lors du Drag, alors...
  ForEach ($drag in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)){
    $_.Effect = 'None'
    #Si l'objet copié est un dossier, alors...
    If((Get-Item -LiteralPath $drag -Force) -is [System.IO.DirectoryInfo]){
      #Détermine le dossier racine du TreeView
      $RacinePath = (Get-Item -LiteralPath $drag -Force)
      #Réinitialise le TreeView
      $treeView.Nodes.Clear()
      #Création de la racine du treeView
      $treeNodes = New-Object System.Windows.Forms.TreeNode
      $treeNodes.text = Split-Path -Path $RacinePath -leaf
      $treeNodes.Name = $treeNodes.text
      $treeNodes.Tag = $treeNodes.text
      $treeView.Nodes.Add($treeNodes) | Out-Null
      $treeView.add_AfterSelect({$textbox.Text = $this.SelectedNode.Name})
      #Création du tronc du treeView
      $array=@(Get-ChildItem -LiteralPath $RacinePath -Force)
      ForEach($item in $array){
        If(!((Get-Item -Path $item.FullName -Force) -is [System.IO.DirectoryInfo])){Add-Node $treeNodes $item | Out-Null}
        #Appel de la fonction de création des branches et feuilles du TreeView pour les dossier
        If((Get-Item -Path $item.FullName -Force) -is [System.IO.DirectoryInfo]){Get-NextLevel $treeNodes $item.FullName}
      }
      #Développe l'arborescence
      $treeNodes.Expand()
      #Compte les dossiers et fichiers contenus dans l'arborescence
      $textboxNbDir.Text = @(Get-ChildItem -LiteralPath $RacinePath -Force -File -Recurse).Count
      $textboxNbFiles.Text = @(Get-ChildItem -LiteralPath $RacinePath -Force -Directory -Recurse).Count + 1
      #Affiche le formulaire treeView 
      $formTree.ShowDialog() | Out-Null
    }
    #Sinon, s'il ne s'agit pas d'un dossier et que l'objet n'est pas présent dans le ListBox, alors on l'y ajoute.
    ElseIf(!($listBox.Items -Contains $drag) -And (!((Get-Item -LiteralPath $drag -Force) -is [System.IO.DirectoryInfo]))){$listBox.Items.Add($drag)}
  }
  If($script:ExitTreeForm){$listBox.Items.Clear(); $script:ExitTreeForm = $False}
  #Réinitialisation la barre de progrossion et mise à jour de la barre de statut
  BarsRefresh
}

$listBox.Add_KeyDown({
  If($_.KeyCode -eq "Delete"){
    $listBox.Items.Remove($listBox.SelectedItem)
    BarsRefresh
  }
})


$RadioButton1.Add_Click({
  BarsRefresh
  $comboBox1.SelectedIndex=0
  $comboBox2.SelectedIndex=6
  $comboBox1.Enabled = $True
  $comboBox2.Enabled = $True
})


$RadioButton2.Add_Click({
  BarsRefresh
  $comboBox1.SelectedIndex=-1
  $comboBox1.Enabled = $False
  $comboBox2.SelectedIndex=-1
  $comboBox2.Enabled = $False
})


$RadioButton3.Add_Click({
  BarsRefresh
  $comboBox3.SelectedIndex=13
  $comboBox3.Enabled = $False
})


$RadioButton4.Add_Click({
  BarsRefresh
  $comboBox3.Enabled = $True
  $comboBox3.SelectedIndex=13
})


$combobox1.Add_SelectionChangeCommitted({
  BarsRefresh
  If($comboBox1.SelectedIndex -eq 1){$comboBox2.SelectedIndex=-1; $comboBox2.Enabled = $False;}
  Else{$comboBox2.Enabled = $True; $comboBox2.SelectedIndex=6}
})


$combobox2.Add_SelectionChangeCommitted({BarsRefresh})


$combobox3.Add_TextChanged({BarsRefresh})


$button_elli.add_click({ 
  BarsRefresh
  #Création de l'objet
  $save = New-Object System.Windows.Forms.SaveFileDialog
  #Initialisation du chemin par défaut.
  $save.initialDirectory = $textBox1.Text.Substring(0, $textBox1.Text.LastIndexOf('\') + 1)
  #$save.RestoreDirectory = $True
  #Titre de la boite de dialogue
  $save.Title = "Enregistrer l'archive sous..."
  #Création d'un filtre sur le type de fichier
  $save.filter = "CAB File (*.cab)| *.cab"
  #Récupération du nom de fichier si présent dans la textBox1
  $save.filename = Split-Path $textBox1.Text -Leaf
  #Affiche la fenêtre d'ouverture de fichier.
  $Sel = $save.ShowDialog()
  #Traitement du retour.
  #Si "OK" on affiche le chemin dans la TextBox.
  #Sinon on afficher un fichier par défaut.
  If($Sel -eq "OK"){
    $SaveAs =  $save.filename.Substring(0, $save.filename.LastIndexOf('\'))
    $textBox1.Text = $save.filename
  }
})


$ButtonTree.Add_Click({
  $script:ExitTreeForm = $False
  If(!($listBox.Items -Contains "$RacinePath\*")){$listBox.Items.Add("$RacinePath\*")}
  $formTree.Close()
  BarsRefresh
})


$ButtonSubFiles.Add_Click({
  $script:ExitTreeForm = $False
  ForEach ($item in (Get-ChildItem -Path $RacinePath -File -Recurse)){If(!($listBox.Items -Contains $item)){$listBox.Items.Add($item.FullName)}}
  $formTree.Close()
  BarsRefresh
})

$ButtonCancel.Add_Click({
  $script:ExitTreeForm = $True
  $formTree.Close()
  BarsRefresh
})


$formTree_FormClosed = ({
  If($script:ExitTreeForm){$listBox.Items.Clear()}
  $formTree.Close()
  BarsRefresh
})

$form_FormClosed = {
  Try{
    [void](Remove-Item $ddf,$WorkingDir -Force -Recurse -ErrorAction SilentlyContinue)
    $listBox.remove_Click($button_ok_Click)
    $listBox.remove_DragOver($listBox_DragOver)
    $listBox.remove_DragDrop($listBox_DragDrop)
    $listBox.remove_DragDrop($listBox_DragDrop)
    $form.remove_FormClosed($Form_Cleanup_FormClosed)
  }
  Catch [Exception]{ }
}


# Attache les évènements à la fenêtre.
$listBox.Add_DragOver($listBox_DragOver)
$listBox.Add_DragDrop($listBox_DragDrop)
$button_addFolder.Add_Click($button_addFolder_Click)
$button_addFile.Add_Click($button_addFile_Click)
$button_supp.Add_Click($button_supp_Click)
$RadioButton1.Add_Click($RadioButton1_CheckedChanged)
$RadioButton2.Add_Click($RadioButton2_CheckedChanged)
$RadioButton3.Add_Click($RadioButton3_CheckedChanged)
$RadioButton4.Add_Click($RadioButton4_CheckedChanged)
$comboBox1.Add_Click($comboBox1_Click)
$comboBox2.Add_Click($comboBox2_Click)
$comboBox3.Add_Click($comboBox3_Click)
$button_ok.Add_Click($button_ok_Click)
$button_quit.Add_Click($button_quit_Click)
$form.Add_FormClosed($form_FormClosed)
$formTree.Add_FormClosed($formTree_FormClosed)
$Button.Add_Click($open_close)
$button1_OnClick = {$formTree.Close()}

# Affichage du formulaire
[void]$form.ShowDialog()

#################################################
# FIN DE PROGRAMME
#################################################
