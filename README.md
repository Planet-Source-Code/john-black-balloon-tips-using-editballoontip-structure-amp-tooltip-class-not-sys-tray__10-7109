<div align="center">

## Balloon Tips using EDITBALLOONTIP structure &amp; ToolTip class \(not sys tray\)


</div>

### Description

Balloon Help Version 2

2 Methods of displaying Balloon Tip Help:

1. Uses EDITBALLOONTIP structure, EM_SHOWBALLOONTIP message and SendMessage to add balloon tip help to your textboxes

2. Uses ToolTip class to add balloon tip help to your other form controls (except ListViews, ListBoxs, TreeViews and RichTextBoxes)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Black](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-black.md)
**Level**          |Intermediate
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB\.NET
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__10-1.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-black-balloon-tips-using-editballoontip-structure-amp-tooltip-class-not-sys-tray__10-7109/archive/master.zip)





### Source Code

<style>
	p {font-family: Courier New; font-size: 12px;}
	.summary {color: rgb(128,128,128);}
	.comment {color: rgb(0,128,0);}
	.keyword {color: rgb(0,0,255);}
	.api {color: rgb(163,21,21);}
</style>
<p>
  <font class="keyword">Imports</font> System.Drawing.SystemColors<br/>
	<font class="keyword">Imports</font> System.Runtime.InteropServices<br/>
	<font class="keyword">Imports</font> System.Windows.Forms<br/>
  <br/>
    <font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
    <font class="comment">''' Class providing method for showing Balloon Tips</font><br/>
    <font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
    <font class="keyword">Public Class</font> BalloonHelp<br/>
      <br/>
			<font class="keyword">#Region</font> <font class="api">" Enumerations: Global "</font><br/>
			<br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
        &nbsp;&nbsp;<font class="comment">''' Balloon Icon Types</font><br/>
        &nbsp;&nbsp;<font class="summary">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
        &nbsp;&nbsp;<font class="keyword">Public Enum</font> BalloonIcon<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;  ShowNone = ToolTipIcon.None<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;  ShowInformation = ToolTipIcon.Info<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;  ShowWarning = ToolTipIcon.Warning<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;  ShowError = ToolTipIcon.Error<br/>
        &nbsp;&nbsp;<font class="keyword">End Enum</font><br/>
			<br/>
			<font class="keyword">#End Region</font><br/>
			<br/>
			<font class="keyword">#Region</font> <font class="api">" Private Declarations: Balloon Help for Textboxes "</font><br/>
			<br/>
  			&nbsp;&nbsp;<font class="keyword">Private Const</font> EM_SHOWBALLOONTIP <font class="keyword">As UInteger</font> = &H1503<br/>
  			<br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
        &nbsp;&nbsp;<font class="comment">''' Type EDITBALLOONTIP converted to a .NET Structure</font><br/>
        &nbsp;&nbsp;<font class="summary">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>	&nbsp;&nbsp;&#60&#83&#116&#114&#117&#99&#116&#76&#97&#121&#111&#117&#116&#40&#76&#97&#121&#111&#117&#116&#75&#105&#110&#100&#46&#83&#101&#113&#117&#101&#110&#116&#105&#97&#108&#41&#62 _<br/>
        &nbsp;&nbsp;<font class="keyword">Private Structure</font> EDITBALLOONTIP<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Public</font> cbStruct <font class="keyword">As Integer</font><br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&#60&#77&#97&#114&#115&#104&#97&#108&#65&#115&#40&#85&#110&#109&#97&#110&#97&#103&#101&#100&#84&#121&#112&#101&#46&#76&#80&#87&#83&#116&#114&#41&#62 <font class="keyword">Public</font> pszTitle <font class="keyword">As String</font><br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&#60&#77&#97&#114&#115&#104&#97&#108&#65&#115&#40&#85&#110&#109&#97&#110&#97&#103&#101&#100&#84&#121&#112&#101&#46&#76&#80&#87&#83&#116&#114&#41&#62 <font class="keyword">Public</font> pszText <font class="keyword">As String</font><br/>
          &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Public</font> ttiIcon <font class="keyword">As Integer</font><br/>
        &nbsp;&nbsp;<font class="keyword">End Structure</font><br/>
  			<br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' Unicode version of SendMessage API needed for pszText + pszText in EDITBALLOONTIP as these are unicode parameters</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
        &nbsp;&nbsp;<font class="keyword">Private Declare Unicode Function</font> SendMessage <font class="keyword">Lib</font> <font class="api">"User32"</font> <font class="keyword">Alias</font> <font class="api">"SendMessageW"</font> (<font class="keyword">ByVal</font> hWnd <font class="keyword">As</font> IntPtr, <font class="keyword">ByVal</font> Msg <font class="keyword">As UInteger</font>, <font class="keyword">ByVal</font> wParam <font class="keyword">As</font> IntPtr, <font class="keyword">ByVal</font> lParam <font class="keyword">As</font> IntPtr) <font class="keyword">As</font> IntPtr<br/>
      <br/>
      <font class="keyword">#End Region</font><br/>
			<br/>
			<font class="keyword">#Region</font> <font class="api">" Private Declarations: Balloon Help / ToolTip for all Controls "</font><br/>
			<br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' Delay times for ToolTips</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
  			&nbsp;&nbsp;<font class="keyword">Private Const</font> DELAY_AUTOPOPUP <font class="keyword">As Integer</font> = 5000<br/>
  			&nbsp;&nbsp;<font class="keyword">Private Const</font> DELAY_INITIAL <font class="keyword">As Integer</font> = 500<br/>
  			&nbsp;&nbsp;<font class="keyword">Private Const</font> DELAY_RESHOW <font class="keyword">As Integer</font> = 500<br/>
			<br/>
      <font class="keyword">#End Region</font><br/>
			<br/>
			<font class="keyword">#Region</font> <font class="api">" Subroutines: Balloon Help for Textboxes "</font><br/>
			<br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' Balloon Help for Textboxes</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;ctlSource&#34&#62;</font><font class="comment">The textbox you want the ToolTip to be displayed under</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;strToolTipText&#34&#62;</font><font class="comment">The ToolTip text</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;strToolTipTitle&#34&#62;</font><font class="comment">The ToolTip caption</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;enmToolTipIcon&#34&#62;</font><font class="comment">The ToolTip icon - None, Info, Warning or Error</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#114&#101&#109&#97&#114&#107&#115&#62<font class="comment">Use this method for textboxes as additional system balloon notifications are displayed and restrictions implemented. If the textbox is a password field, the user will be prompted if the capslock is on and will also not be able to cut text</font></font><font class="summary">&#60&#47&#114&#101&#109&#97&#114&#107&#115&#62;</font><br/>
        &nbsp;&nbsp;<font class="keyword">Public Shared Sub</font> Show(<font class="keyword">ByVal</font> ctlSource <font class="keyword">As</font> Control, <font class="keyword">ByVal</font> strToolTipText <font class="keyword">As String</font>, <font class="keyword">ByVal</font> strToolTipTitle <font class="keyword">As String</font>, <font class="keyword">Optional ByVal</font> enmToolTipIcon <font class="keyword">As</font> BalloonIcon = BalloonIcon.ShowInformation)<br/>
  				<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Dim</font> EBT <font class="keyword">As</font> EDITBALLOONTIP = <font class="keyword">New</font> EDITBALLOONTIP<br/>
  				<br/>
            &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">With</font> EBT<br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  .cbStruct = Marshal.SizeOf(EBT)<br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  .ttiIcon = enmToolTipIcon<br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  .pszText = strToolTipText<br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  .pszTitle = strToolTipTitle<br/>
            &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">End With</font><br/>
  				<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Dim</font> ptrEBT <font class="keyword">As</font> IntPtr = Marshal.AllocHGlobal(EBT.cbStruct)<br/>
  				<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;Marshal.StructureToPtr(EBT, ptrEBT, False)<br/>
  				<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Call</font> SendMessage(ctlSource.Handle, EM_SHOWBALLOONTIP, IntPtr.Zero, ptrEBT)<br/>
  				<br/>
        &nbsp;&nbsp;<font class="keyword">End Sub</font><br/>
			<br/>
			<font class="keyword">#End Region</font><br/>
			<br/>
			<font class="keyword">#Region</font> <font class="api">" Subroutines: Balloon Help / ToolTip for all Controls "</font><br/>
			<br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' Balloon Help / ToolTip for all Controls</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      	&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;ctlSource&#34&#62;</font><font class="comment">The control you want the ToolTip to be displayed under</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;strToolTipText&#34&#62;</font><font class="comment">The ToolTip text</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;strToolTipTitle&#34&#62;</font><font class="comment">The ToolTip caption</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;clrBackColor&#34&#62;</font><font class="comment">The backcolor of the ToolTip (any System.Drawing.Color)</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;clrForeColor&#34&#62;</font><font class="comment">The text colour of the ToolTip (any System.Drawing.Color)</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;enmToolTipIcon&#34&#62;</font><font class="comment">The ToolTip icon - None, Info, Warning or Error</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;intAutoPopDelay&#34&#62;</font><font class="comment">The period of time the ToolTip remains visible on control mousehover event</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;intInitialDelay&#34&#62;</font><font class="comment">The period of time before the ToolTip appears with mouseover event</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;intReshowDelay&#34&#62;</font><font class="comment">The period of time before subsequent ToolTip windows appear</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;blnIsBalloon&#34&#62;</font><font class="comment">Show as a balloon tip (True) or a standard tooltip (False)</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;blnUseAnimation&#34&#62;</font><font class="comment">Use animation - Only XP, Windows Server 2003, IE 5+</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;blnUseFading&#34&#62;</font><font class="comment">Use fading - Only XP, Windows Server 2003, IE 5+</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;blnActive&#34&#62;</font><font class="comment">Enable the tooltip</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#112&#97&#114&#97&#109&#32&#110&#97&#109&#101&#61&#34;blnShowAlways&#34&#62;</font><font class="comment">Force the ToolTip text to be displayed regardless if form has focus or not</font><font class="summary">&#60&#47&#112&#97&#114&#97&#109&#62;</font><br/>
  			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#114&#101&#109&#97&#114&#107&#115&#62<font class="comment">This method can be used with all form controls except ListViews, ListBoxs, TreeViews and RichTextBoxes</font></font><font class="summary">&#60&#47&#114&#101&#109&#97&#114&#107&#115&#62;</font><br/>
  			&nbsp;&nbsp;<font class="keyword">Public Sub</font> Show(<font class="keyword">ByVal</font> ctlSource <font class="keyword">As</font> Control, <font class="keyword">ByVal</font> strToolTipText <font class="keyword">As String</font>, <font class="keyword">ByVal</font> strToolTipTitle <font class="keyword">As String</font>, <font class="keyword">ByVal</font> clrBackColor <font class="keyword">As</font> Color, <font class="keyword">ByVal</font> clrForeColor <font class="keyword">As</font> Color, <font class="keyword">Optional ByVal</font> enmToolTipIcon <font class="keyword">As</font> BalloonIcon = BalloonIcon.ShowNone, <font class="keyword">Optional ByVal</font> intAutoPopDelay <font class="keyword">As Integer</font> = DELAY_AUTOPOPUP, <font class="keyword">Optional ByVal</font> intInitialDelay <font class="keyword">As Integer</font> = DELAY_INITIAL, <font class="keyword">Optional ByVal</font> intReshowDelay <font class="keyword">As Integer</font> = DELAY_RESHOW, <font class="keyword">Optional ByVal</font> blnIsBalloon <font class="keyword">As Boolean</font> = <font class="keyword">True</font>, <font class="keyword">Optional ByVal</font> blnUseAnimation <font class="keyword">As Boolean</font> = <font class="keyword">True</font>, <font class="keyword">Optional ByVal</font> blnUseFading <font class="keyword">As Boolean</font> = <font class="keyword">True</font>, <font class="keyword">Optional ByVal</font> blnActive <font class="keyword">As Boolean</font> = <font class="keyword">True</font>, <font class="keyword">Optional ByVal</font> blnShowAlways <font class="keyword">As Boolean</font> = <font class="keyword">False</font>)<br/>
  		<br/>
  		&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Dim</font> MyToolTip <font class="keyword">As New</font> ToolTip()<br/>
  		<br/>
  		&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">With</font> MyToolTip<br/>
  		<br/>
  			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="comment">' Set up the delays for the ToolTip.</font><br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.AutoPopDelay = intAutoPopDelay<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.InitialDelay = intInitialDelay<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ReshowDelay = intReshowDelay<br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="comment">' Set up the appearance of the ToolTip</font><br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.BackColor = clrBackColor<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ForeColor = clrForeColor<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.IsBalloon = blnIsBalloon<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ToolTipIcon = enmToolTipIcon<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.UseAnimation = blnUseAnimation<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.UseFading = blnUseFading<br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="comment">' Set up the ToolTip display options</font><br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.Active = blnActive<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ShowAlways = blnShowAlways<br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="comment">' Set ToolTip Caption and Text</font><br/>
    		<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ToolTipTitle = strToolTipTitle<br/>
    		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.SetToolTip(ctlSource, strToolTipText)<br/>
    		<br/>
  		&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">End With</font><br/>
  		<br/>
  		&nbsp;&nbsp;<font class="keyword">End Sub</font><br/>
		<br/>
		<font class="keyword">#End Region</font><br/>
		<br/>
		<font class="keyword">#Region</font> <font class="api">" Properties: Balloon Help / ToolTip for all Controls "</font><br/>
		<br/>
			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      &nbsp;&nbsp;<font class="comment">''' System defined ToolTip background colour</font><br/>
  		&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#114&#101&#116&#117&#114&#110&#115&#62<font class="comment">System.Drawing.SystemColors.Info</font></font><font class="summary">&#60&#47&#114&#101&#116&#117&#114&#110&#115&#62;</font><br/>
			&nbsp;&nbsp;<font class="keyword">Public ReadOnly Property</font> TipBackColour() <font class="keyword">As</font> Color<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Get</font><br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Return</font> Info<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">End Get</font><br/>
      &nbsp;&nbsp;<font class="keyword">End Property</font><br/>
		<br/>
			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
      &nbsp;&nbsp;<font class="comment">''' System defined ToolTip text colour</font><br/>
  		&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#47&#115&#117&#109&#109&#97&#114&#121&#62;</font><br/>
			&nbsp;&nbsp;<font class="comment">''' </font><font class="summary">&#60&#114&#101&#116&#117&#114&#110&#115&#62<font class="comment">System.Drawing.SystemColors.InfoText</font></font><font class="summary">&#60&#47&#114&#101&#116&#117&#114&#110&#115&#62;</font><br/>
			&nbsp;&nbsp;<font class="keyword">Public ReadOnly Property</font> TipTextColour() <font class="keyword">As</font> Color<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Get</font><br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Return</font> InfoText<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">End Get</font><br/>
      &nbsp;&nbsp;<font class="keyword">End Property</font><br/>
		<br/>
		<font class="keyword">#End Region</font><br/>
		<br/>
    <font class="keyword">End Class</font><br/>
	<br/>
  <font class="comment"><b><u>Usage:</b><br/>
	<br/>
  For a Textbox Information Balloon tip...</u></font><br/>
	<br/>
    &nbsp;&nbsp;<font class="keyword">Call</font> BalloonHelp.Show(myTextbox, Message, Caption, BalloonHelp.BalloonIcon.ShowInformation)<br/>
	<br/>
	&nbsp;&nbsp;Where <font class="api">myTextbox</font> is the control you want the ToolTip to be displayed under,<br/>
	&nbsp;&nbsp;Where <font class="api">Message</font> is the ToolTip Message text,<br/>
	&nbsp;&nbsp;Where <font class="api">Caption</font> is the ToolTip Title text,<br/>
	&nbsp;&nbsp;Where <font class="api">BalloonHelp.BalloonIcon</font> is the ToolTip icon (None, Information, Warning or Error).<br/>
	<br/>
  <font class="comment"><u>For a Button Information Balloon tip...</u></font><br/>
	<br/>
		&nbsp;&nbsp;<font class="keyword">Dim</font> ttpButton <font class="keyword">As New</font> BalloonHelp<br/>
		<br/>
    &nbsp;&nbsp;<font class="keyword">With</font> ttpButton<br/>
      &nbsp;&nbsp;&nbsp;&nbsp;.Show(myButton, Message, Caption, .TipBackColour, .TipTextColour, BalloonHelp.BalloonIcon.ShowInformation)<br/>
    &nbsp;&nbsp;<font class="keyword">End With</font><br/>
	<br/>
	&nbsp;&nbsp;Where <font class="api">myButton</font> is the Button you want the ToolTip to be displayed above,<br/>
	&nbsp;&nbsp;Where <font class="api">Message</font> is the ToolTip Message text,<br/>
	&nbsp;&nbsp;Where <font class="api">Caption</font> is the ToolTip Title text,<br/>
	&nbsp;&nbsp;Where <font class="api">.TipBackColour</font> is the system default ToolTip background colour (BalloonHelp.TipBackColour property),<br/>
	&nbsp;&nbsp;Where <font class="api">.TipTextColour</font> is the system default ToolTip text colour (BalloonHelp.TipTextColour property),<br/>
	&nbsp;&nbsp;Where <font class="api">BalloonHelp.BalloonIcon</font> is the ToolTip icon (None, Information, Warning or Error).<br/>
	<br/>
  <font class="comment"><u>For a customised PictureBox Information Balloon tip with a black background and white text...</u></font><br/>
	<br/>
		&nbsp;&nbsp;<font class="keyword">Imports</font> System.Drawing.Color<br/>
		<br/>
		&nbsp;&nbsp;...<br/>
		<br/>
		&nbsp;&nbsp;<font class="keyword">Dim</font> ttpPictureBox <font class="keyword">As New</font> BalloonHelp<br/>
		<br/>
    &nbsp;&nbsp;ttpPictureBox.Show(myPictureBox, Message, Caption, Black, White, BalloonHelp.BalloonIcon.ShowInformation)<br/>
	<br/>
	&nbsp;&nbsp;Where <font class="api">myPictureBox</font> is the PictureBox you want the ToolTip to be displayed above,<br/>
	&nbsp;&nbsp;Where <font class="api">Message</font> is the ToolTip Message text,<br/>
	&nbsp;&nbsp;Where <font class="api">Caption</font> is the ToolTip Title text,<br/>
	&nbsp;&nbsp;Where <font class="api">Black</font> is the ToolTip background colour = System.Drawing.Color.Black,<br/>
	&nbsp;&nbsp;Where <font class="api">White</font> is the ToolTip background colour = System.Drawing.Color.White,<br/>
	&nbsp;&nbsp;Where <font class="api">BalloonHelp.BalloonIcon</font> is the ToolTip icon (None, Information, Warning or Error).<br/>
</p>

