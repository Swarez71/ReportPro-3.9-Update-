unit RpFormUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Buttons, StdCtrls, OleCtrls, RpRuntime_TLB, ComCtrls, ExtCtrls;

type

   // Record and pointer type for accessing treeview data.
   PTreeItemData = ^TTreeItemData;
   TTreeItemData = record
      ItemID  : integer;
      Section : integer;
      Name    : string;
   end;

   // Record and pointer type for accessing listview data.
   PListItemData = ^TListItemData;
   TListItemData = record
      Attribute  : integer;
      EditOption : integer;
      Cargo      : variant;
   end;

  TfrmReport = class(TForm)
    pnlRpName: TPanel;
    edRpName: TEdit;
    spRpName: TSpeedButton;
    pnlRpRun: TPanel;
    tvRpRun: TTreeView;
    lvRpRun: TListView;
    pnlButtons: TPanel;
    dlgRpOpen: TOpenDialog;
    pbPrintDialog: TButton;
    pbSetupDialog: TButton;
    pbPreview: TButton;
    pbPrint: TButton;
    pbExport: TButton;
    pbClose: TButton;
    splRpRun: TSplitter;
    pbExpression: TButton;
    ckEvents: TCheckBox;
    lblRpName: TLabel;
    xRpRun: TRpRuntime;
    procedure spRpNameClick(Sender: TObject);
    procedure pbCloseClick(Sender: TObject);
    procedure pbExportClick(Sender: TObject);
    procedure pbPrintClick(Sender: TObject);
    procedure pbPreviewClick(Sender: TObject);
    procedure tvRpRunChange(Sender: TObject; Node: TTreeNode);
    procedure lvRpRunChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure pbExpressionClick(Sender: TObject);
    procedure xRpRunReportStart(Sender: TObject);
    procedure xRpRunReportEnd(Sender: TObject);
    procedure pbSetupDialogClick(Sender: TObject);
    procedure pbPrintDialogClick(Sender: TObject);
  private

    // Methods for loading information into the treeview.
    procedure AddChildrenTables2TreeView( nSection : integer; oParent : TTreeNode; sTable : string );
    function AddItem2TreeView( oParent : TTreeNode; sCaption : string; nItemID, nSection : integer; sName : string ) : TTreeNode;
    procedure LoadTreeView;

    // Methods for loading information into the listview.
    procedure AddItem2ListView( sCaption, sData : string; nAttribute, nEditOption : integer; vCargo : variant );
    procedure LoadListView( oTreeNode : TTreeNode );
    procedure LoadPrinterInfo;
    procedure LoadReportInfo;
    procedure LoadSectionInfo( pData : PTreeItemData );
    procedure LoadTableInfo( pData : PTreeItemData );
    procedure LoadVariableInfo( pData : PTreeItemData );

    // Methods for enabling and clearing the controls.
    procedure ClearEdits;
    procedure EnableControls;

    // Methods for releasing the data stored in the treeview and listview items.
    procedure ReleaseTreeViewData;
    procedure ReleaseListViewData;

    // Methods for setting the attibutes of the report.
    procedure SetIntAttribute( nItemID, nSection : integer; sName : string; nAttribute, nValue : integer );
    procedure SetStringAttribute( nItemID, nSection : integer; sName : string; nAttribute : integer; sValue : string );

  public
  end;

var
  frmReport: TfrmReport;

implementation

{$R *.DFM}

const
   EO_NEVER       = -1; // Edit option: None
   EO_EXPBUILDER  = 01; // Edit option: Show Expression builder
   EO_INPLACEEDIT = 02; // Edit option: Allow in-place editing
   EO_LOGIC       = 03; // Edit option: Logic toggle
   EO_POPUPEDIT   = 04; // Edit option: Show data in popup dialog
   EO_SPINNER     = 05; // Edit option: spinner
   LVI_ATTRIBUTE  = 01; // ListViewItem:Value array element: The attribute we are displaying
   LVI_EDITOPTION = 02; // ListViewItem:Value array element: One of the EO_ defines.
   LVI_CARGO      = 03; // ListViewItem:Value array element: Cargo for misc edit options
   RIT_REPORT     = 01; // Report Info Type: The type of information in the list view
   RIT_PRINTER    = 02; // Report Info Type: The type of information in the list view
   RIT_SECTION    = 03; // Report Info Type: The type of information in the list view
   RIT_TABLE      = 04; // Report Info Type: The type of information in the list view
   RIT_VARIABLE   = 05; // Report Info Type: The type of information in the list view
   TVI_ITEMID     = 01; // TreeViewItem:Value array element: One of the RIT_ defines.
   TVI_SECTION    = 02; // TreeViewItem:Value array element: The section number we are looking at
   TVI_NAME       = 03; // TreeViewItem:Value array element: Used to store table or variable names

{******************************************************************************}
{************************ Button Event Handling Methods ***********************}
{******************************************************************************}

procedure TfrmReport.spRpNameClick(Sender: TObject);
begin

   // If the report is not closed...
   if ( not self.xRpRun.Close ) then begin

      // There are rare cases where you can end up here
      // and the report is cannot properly shut down.
      // In that case, display an error message.
      raise Exception.Create( 'The current report is busy and cannot be closed at this time.#13#10Please try again later.' );

   end;

   // Clear the data out of the controls.
   self.ClearEdits;

   // If the user pressed the okay button on the open dialog...
   if ( self.dlgRpOpen.Execute ) then begin

      // If they actually selected a valid file name...
      if ( self.dlgRpOpen.FileName <> '' ) then begin

         // Get the file name selected by the user.
         //self.edRpName.Text := self.dlgRpOpen.FileName;

         // Load the report selected by the user.
         self.xRpRun.LoadReport( self.dlgRpOpen.FileName );

         // If the report loaded correctly...
         if ( self.xRpRun.IsValid ) then begin

            // Even though we already have the file name, we'll do it long hand
            // to show how to retrieve the file name from the report.
            self.edRpName.Text := self.xRpRun.GetReportStringAttribute( RPT_ATTR_REPORT_FILE );

            // Load the data from the report into the treeview.
            self.LoadTreeView;

         end

         else begin

            // Close the report.
            self.xRpRun.Close;

            // Display an error message.
            raise Exception.Create( 'An error occurred while opening the report.' );

         end;

      end;

   end;

   // Enable the controls if the report opened correctly.
   self.EnableControls;

end;

procedure TfrmReport.pbExpressionClick(Sender: TObject);
var
   oTVItem : TTreeNode;     // Reference to the selected treeview item
   oLVItem : TListItem;     // Reference to the selected listview item
   pTVData : PTreeItemData; // Pointer to the data in the selected treeview item
   pLVData : PListItemData; // Pointer to the data in the selected listview item
   sResult : string;        // Resulting expression from the expression builder

begin

   // Retrieve the selected items from the treeview and listview controls.
   oTVItem := self.tvRpRun.Selected;
   oLVItem := self.lvRpRun.Selected;
   pTVData := oTVItem.Data;
   pLVData := oLVItem.Data;

   // Show the expression builder.
   sResult := self.xRpRun.ExpressionBuilder( 00, pTVData^.Section, oLVItem.Caption,
                                             Trim( oLVItem.SubItems[ 00 ] ), true,
                                             pLVData^.Cargo, true, true, true );

   // Write the expression result to the listview.
   oLVItem.SubItems[ 00 ] := sResult;

   // Write the expression result to the report.
   self.SetStringAttribute( pTVData^.ItemID, pTVData^.Section, pTVData^.Name,
                            pLVData^.Attribute, sResult );

end;

{
   This method is called when the print dialog button is pressed.
}
procedure TfrmReport.pbPrintDialogClick(Sender: TObject);
var
   oTVItem : TTreeNode;     // Reference to the selected treeview item
   pTVData : PTreeItemData; // Pointer to the data in the selected treeview item

begin

   // If the user changed the data in the printer dialog...
   if ( self.xRpRun.ShowPrintDlg( UPDATE_PAPERINFO_PROMPTUSER ) ) then begin

      // Retrieve the selected items from the treeview control.
      oTVItem := self.tvRpRun.Selected;
      pTVData := oTVItem.Data;

      // If we are displaying printer attributes, refresh the grid...
      if ( pTVData^.ItemID = RIT_PRINTER ) then begin

         // Release the data stored in the listview items.
         self.ReleaseListViewData;

         // Clear the items from the listview.
         self.lvRpRun.Items.Clear;

         // Load the new printer information.
         self.LoadPrinterInfo;

      end;

   end;

end;

{
   This method is called when the print setup dialog button is pressed.
}
procedure TfrmReport.pbSetupDialogClick(Sender: TObject);
var
   oTVItem : TTreeNode;     // Reference to the selected treeview item
   pTVData : PTreeItemData; // Pointer to the data in the selected treeview item

begin

   // If the user changed the data in the printer setup dialog...
   if ( self.xRpRun.ShowPrinterSetupDlg( UPDATE_PAPERINFO_PROMPTUSER ) ) then begin

      // Retrieve the selected items from the treeview control.
      oTVItem := self.tvRpRun.Selected;
      pTVData := oTVItem.Data;

      // If we are displaying printer attributes, refresh the grid...
      if ( pTVData^.ItemID = RIT_PRINTER ) then begin

         // Release the data stored in the listview items.
         self.ReleaseListViewData;

         // Clear the items from the listview.
         self.lvRpRun.Items.Clear;

         // Load the new printer information.
         self.LoadPrinterInfo;

      end;

   end;

end;

procedure TfrmReport.pbPreviewClick(Sender: TObject);
begin

   // Start the print preview process.
   self.xRpRun.PreviewReport;

end;

procedure TfrmReport.pbPrintClick(Sender: TObject);
begin

   // Start the print process.
   self.xRpRun.PrintReport;

end;

procedure TfrmReport.pbExportClick(Sender: TObject);
begin

   // Start the export process.
   self.xRpRun.ExportReport;

end;

procedure TfrmReport.pbCloseClick(Sender: TObject);
begin

   // If the report is not closed...
   if ( not self.xRpRun.Close ) then begin

      // There are rare cases where you can end up here
      // and the report is cannot properly shut down.
      // In that case, display an error message.
      raise Exception.Create( 'The current report is busy and cannot be closed at this time.#13#10Please try again later.' );

   end;

   // Close the form.
   self.Close;

end;

{******************************************************************************}
{*********************** TreeView Event Handling Methods **********************}
{******************************************************************************}

{
   This method is called when the selection in the tree view changes.
}
procedure TfrmReport.tvRpRunChange(Sender: TObject; Node: TTreeNode);
begin

   // Load the listview with information for the current treeview item.
   self.LoadListView( Node );

   // Disable the expression builder button
   // until the user selects a listview item.
   self.pbExpression.Enabled := false;

end;

{******************************************************************************}
{*********************** ListView Event Handling Methods **********************}
{******************************************************************************}

procedure TfrmReport.lvRpRunChange(Sender: TObject; Item: TListItem; Change: TItemChange);
var
   pData : PListItemData; // Pointer to the data in the selected listview item

begin

   // Get a reference to the data in the current listview item.
   pData := Item.Data;

   // If data was assigned to the current listview item...
   if ( Assigned( pData ) ) then begin

      // If the expression builder is available for this item...
      if ( pData^.EditOption = EO_EXPBUILDER ) then begin

         // Enable the expression builder button.
         self.pbExpression.Enabled := true;

      end

      else begin

         // Disable the expression builder button.
         self.pbExpression.Enabled := false;

      end;

   end;

end;

{******************************************************************************}
{*********************** ActiveX Event Handling Methods ***********************}
{******************************************************************************}

procedure TfrmReport.xRpRunReportStart(Sender: TObject);
begin

   // If the user wants to see the events displayed...
   if ( self.ckEvents.Checked ) then begin

      // Display the event to the user.
      MessageDlg( 'The OnReportStart() event was fired.',
                  mtInformation, [ mbOk ], 00 );

   end;

end;

procedure TfrmReport.xRpRunReportEnd(Sender: TObject);
var
   sText : string; // Message text to display to the user

begin

   // If the user wants to see the events displayed...
   if ( self.ckEvents.Checked ) then begin

      // Build the event message string.
      sText := 'The OnReportEnd() event was fired.'#13#10#13#10;

      // If the report completed successfully...
      if ( self.xRpRun.SuccessfulCompletion ) then begin

         // Add the success message to the text.
         sText := sText + 'The report successfully completed.';

      end

      else begin

         // Add the failure message to the text.
         sText := sText + 'The report was terminated by the user.';

      end;

      // Display the event to the user.
      MessageDlg( sText, mtInformation, [ mbOk ], 00 );

   end;

end;

{******************************************************************************}
{*************************** TreeView Fill Methods ****************************}
{******************************************************************************}

{
   This method gathers information about the children of a table.
   Since relationship between tables is hierarchical, this method
   will be called recursively as it adds the tables to our tree
   view control.

   Note: the child table names are returned in a comma delimited string.

   Parameters:

      nSection - The section number we are working in.
      oParent  - The TreeviewItem that will be the parent for each item added to the tree view
      sTable   - The name of the table that we are retrieving children table information for
}
procedure TfrmReport.AddChildrenTables2TreeView( nSection : integer; oParent : TTreeNode; sTable : string );
var
   oItem1, oItem2 : TTreeNode; // Reference to the current treeview item
   sChild, sName  : string;    // Name of the current child table
   nPos           : integer;   // Position of the current delimiter

begin

   // See if this is an SQL Query.
   sChild := Trim( self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_CHILD_SQLTABLES ) );

   // If it is, it will have child SQL tables otherwise this call will fail...
   if ( sChild <> '' ) then begin

      // Add the "SQL Tables" heading.
      oItem1 := self.AddItem2TreeView( oParent, 'SQL Tables', 00, 00, '' );

      // While there are child tables in the string...
      while ( sChild <> '' ) do begin

         // Find the postition of the next delimiter.
         nPos := Pos( ', ', sChild );

         // If the delimiter was found...
         if ( nPos = 00 ) then begin

            // Assign the table names.
            sName  := sChild;
            sChild := '';

         end

         else begin

            // Parse out each of the table names.
            sName  := Copy( sChild, 01, Pred( nPos ) );
            sChild := Copy( sChild, Succ( nPos ), Length( sChild ) - nPos );

         end;

         // Add each of the child SQL tables to the tree view.
         oItem2 := self.AddItem2TreeView( oItem1, sName, RIT_TABLE, nSection, sName );

         // See if this table has child tables.
         self.AddChildrenTables2TreeView( nSection, oItem2, sName );

      end;

   end;

   // Get the child tables string from the report.
   sChild := Trim( self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_CHILD_TABLES ) );

   // Now add the NON SQL tables (DBFs, SQL Queries and Jasmine Queries).
   if ( sChild <> '' ) then begin

      // While there are child tables in the string...
      while ( sChild <> '' ) do begin

         // Find the postition of the next delimiter.
         nPos := Pos( ', ', sChild );

         // If the delimiter was found...
         if ( nPos = 00 ) then begin

            // Assign the table names.
            sName  := sChild;
            sChild := '';

         end

         else begin

            // Parse out each of the table names.
            sName  := Copy( sChild, 01, Pred( nPos ) );
            sChild := Copy( sChild, Succ( nPos ), Length( sChild ) - nPos );

         end;

         // Add each of the child tables to the tree view.
         oItem2 := self.AddItem2TreeView( oParent, sName, RIT_TABLE, nSection, sName );

         // See if this table has child tables.
         self.AddChildrenTables2TreeView( nSection, oItem2, sName );

      end;

   end;

end;

{
   This is a simple method that adds items to the tree view.
   It provides an easy way to have the compiler check the
   data type of the information added to the tree view and
   also reduces the amount of in-line code.

   Parameters:

      oParent  - The parent tree view item for the added item
      sCaption - The displayed text
      nItemID  - The category of data that this item represents (on of the RIT_ defines)
      nSection - The section number where this item resides (if applicable)
      sName    - The name of this item (if applicable)
}
function TfrmReport.AddItem2TreeView( oParent : TTreeNode; sCaption : string; nItemID, nSection : integer; sName : string ) : TTreeNode;
var
   oItem : TTreeNode;     // Reference to the selected treeview item
   pData : PTreeItemData; // Pointer to the data in the selected treeview item

begin

   // Allocate the pointer to the data structure.
   New( pData );

   // Fill in the information.
   pData^.ItemID  := nItemID;
   pData^.Section := nSection;
   pData^.Name    := sName;

   // Add a new child item to the treeview.
   oItem := self.tvRpRun.Items.AddChild( oParent, sCaption );

   // Store the data for the treeview item.
   oItem.Data := pData;

   // Return the item so that subitems can be added to it.
   result := oItem;

end;

{
   This method loads the tree view control with all
   the different report categories and entity names.
}
procedure TfrmReport.LoadTreeView;
var
   oItem1, oItem2, oItem3, oItem4 : TTreeNode; // Reference to the current treeview item
   nSections, nSection            : integer;   // Number of sections in the report
   nVariables, nVariable          : integer;   // Number of variables in the report
   sValue                         : string;    // Current attribute value from the report

begin

   // Add a category for the report attributes.
   self.AddItem2TreeView( nil, 'Report', RIT_REPORT, 00, '' );

   // Add a category for the printer attributes.
   self.AddItem2TreeView( nil, 'Printer', RIT_PRINTER, 00, '' );

   // Add a category for the section attributes and save the
   // section treeview item so we can add children to it.
   oItem1 := self.AddItem2TreeView( nil, 'Sections', 00, 00, '' );

   // Get the number of sections in the report.
   nSections := self.xRpRun.GetReportIntAttribute( RPT_ATTR_SECTION_COUNT );

   // Add a treeview item for each section
   for nSection := 01 to nSections do begin

      // Add a label for each section.
      oItem2 := self.AddItem2TreeView( oItem1, 'Section ' + IntToStr( nSection ), RIT_SECTION, nSection, '' );

      // Add a sub heading for the tables.
      oItem3 := self.AddItem2TreeView( oItem2, 'Tables', 00, 00, '' );

      // Get the primary table for this section.
      sValue := Trim( self.xRpRun.GetSectionStringAttribute( nSection, SECTION_ATTR_PRIMARY_TABLE ) );

      // If there is a table...
      if ( sValue <> '' ) then begin

         // Add the primary table.
         oItem4 := self.AddItem2TreeView( oItem3, sValue, RIT_TABLE, nSection, sValue );

         // Add the children for this table.
         self.AddChildrenTables2TreeView( nSection, oItem4, sValue );

      end;

      // Add a sub heading for the variables.
      oItem3 := self.AddItem2TreeView( oItem2, 'Variables', 00, 00, '' );

      // Add a sub item for each variable.
      nVariables := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_VARIABLE_COUNT );

      for nVariable := 01 to nVariables do begin

         // Get the name of each variable.
         sValue := self.xRpRun.GetVariableStringAttribute( nSection, IntToStr( nVariable ), VARIABLE_ATTR_NAME );

         // Add the variable to the treeview.
         self.AddItem2TreeView( oItem3, sValue, RIT_VARIABLE, nSection, sValue );

      end;

   end;

   // Select the first item in the tree view.
   self.tvRpRun.Selected := oItem1;

   // Load the listview for the first item.
   self.LoadListView( oItem1 );

end;

{******************************************************************************}
{*************************** ListView Fill Methods ****************************}
{******************************************************************************}

{
   This is a simple method that adds items to the grid.
   It provides an easy way to have the compiler check
   the data type of the information added to the grid
   and also reduces the amount of in-line code.

   Parameters:

      sCaption    - The text displayed in the first column of the grid
      sData       - The text displayed in the second column of the grid
      nAttribute  - The attribute ID of the item (one of the report attribute defines)
      nEditOption - The edit option for this attribute (one of the EO_ defines)
      vCargo      - a user defined value depending on the edit option
}
procedure TfrmReport.AddItem2ListView( sCaption, sData : string; nAttribute, nEditOption : integer; vCargo : variant );
var
   oItem : TListItem;     // Reference to the selected listview item
   pData : PListItemData; // Pointer to the data in the selected listview item

begin

   // Allocate the pointer to the data structure.
   New( pData );

   // Fill in the information.
   pData^.Attribute  := nAttribute;
   pData^.EditOption := nEditOption;
   pData^.Cargo      := vCargo;

   // Add a new item to the listview.
   oItem := self.lvRpRun.Items.Add;

   // Assign the caption to display.
   oItem.Caption := sCaption;

   // Add the text for the description.
   oItem.SubItems.Add( sData );

   // Store the data for the listview item.
   oItem.Data := pData;

end;

{
   This method is called when a tree view item is selected.
   It determines what category of information we are viewing
   and loads the appropriate information in the grid.
}
procedure TfrmReport.LoadListView( oTreeNode : TTreeNode );
var
   pData : PTreeItemData; // Pointer to the data in the selected treeview item
   nType : integer;       // Type of information to load

begin

   // Release the data stored in the listview items.
   self.ReleaseListViewData;

   // Delete all the items in the grid.
   self.lvRpRun.Items.Clear;

   // Get the extra information we saved on the tree view item.
   pData := oTreeNode.Data;
   nType := pData^.ItemID;

   // Determine what type of information we want to load.
   case ( nType ) of
      RIT_REPORT   : self.LoadReportInfo;
      RIT_PRINTER  : self.LoadPrinterInfo;
      RIT_SECTION  : self.LoadSectionInfo( pData );
      RIT_TABLE    : self.LoadTableInfo( pData );
      RIT_VARIABLE : self.LoadVariableInfo( pData );
   end;

end;

{
   This method retrieves printer related information
   from the report and loads it into the grid.  It
   also associates the editing options available for
   each attribute.
}
procedure TfrmReport.LoadPrinterInfo;
var
   sValue : string;  // String attribute data from the report
   nValue : integer; // Numeric attribute data from the report

begin

   // Get the attribute information from the report and fill in the listview.
   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PRINTER_NAME );
   self.AddItem2ListView( 'Printer', sValue, RPT_ATTR_PRINTER_NAME, EO_NEVER, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PRINT_JOB_TITLE );
   self.AddItem2ListView( 'Print Job Title', sValue, RPT_ATTR_PRINT_JOB_TITLE, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PRINT2FILE_NAME );
   self.AddItem2ListView( 'Print To File Name', sValue, RPT_ATTR_PRINT2FILE_NAME, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PRINT_CAPTION );
   self.AddItem2ListView( 'Printing Dialog Caption', sValue, RPT_ATTR_PRINT_CAPTION, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PRINT_MESSAGE1 );
   self.AddItem2ListView( 'Printing Dialog Message 1', sValue, RPT_ATTR_PRINT_MESSAGE1, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PRINT_MESSAGE2 );
   self.AddItem2ListView( 'Printing Dialog Message 2', sValue, RPT_ATTR_PRINT_MESSAGE2, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PREVIEW_MODAL );
   self.AddItem2ListView( 'Modal Preview', sValue, RPT_ATTR_PREVIEW_MODAL, EO_LOGIC, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PREVIEW_CAPTION );
   self.AddItem2ListView( 'Preview Caption', sValue, RPT_ATTR_PREVIEW_CAPTION, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_PREVIEW_NOZOOM );
   self.AddItem2ListView( 'Preview No Zoom', sValue, RPT_ATTR_PREVIEW_NOZOOM, EO_LOGIC, 00 );

   nValue := self.xRpRun.GetReportIntAttribute( RPT_ATTR_PREVIEW_ZOOM_MODE );
   self.AddItem2ListView( 'Preview Zoom Mode', IntToStr( nValue ), RPT_ATTR_PREVIEW_ZOOM_MODE, EO_SPINNER, VarArrayOf( [ 00, 02 ] ) );

   nValue := self.xRpRun.GetReportIntAttribute( RPT_ATTR_PREVIEW_PAGECOUNT );
   self.AddItem2ListView( 'Preview Panes', IntToStr( nValue ), RPT_ATTR_PREVIEW_PAGECOUNT, EO_SPINNER, VarArrayOf( [ 01, 02 ] ) );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_EXPORT_FILE_NAME );
   self.AddItem2ListView( 'Export File Name', sValue, RPT_ATTR_EXPORT_FILE_NAME, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_EXPORT_CAPTION );
   self.AddItem2ListView( 'Export Dialog Caption', sValue, RPT_ATTR_EXPORT_CAPTION, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_EXPORT_MESSAGE );
   self.AddItem2ListView( 'Export Dialog Message', sValue, RPT_ATTR_EXPORT_MESSAGE, EO_INPLACEEDIT, 00 );

end;

{
   This method retrieves report related information
   from the report and loads it into the grid.  It
   also associates the editing options available for
   each attribute.
}
procedure TfrmReport.LoadReportInfo;
var
   sValue : string; // String attribute data from the report

begin

   // Get the attribute information from the report and fill in the listview.
   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_REPORT_TITLE );
   self.AddItem2ListView( 'Title', sValue, RPT_ATTR_REPORT_TITLE, EO_NEVER, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_REPORT_DESCRIPTION );
   self.AddItem2ListView( 'Description', sValue, RPT_ATTR_REPORT_DESCRIPTION, EO_POPUPEDIT, true );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_CONNECTED );
   self.AddItem2ListView( 'Connected to datasource(s)', sValue, RPT_ATTR_CONNECTED, EO_NEVER, 00 );

   sValue := self.xRpRun.GetReportStringAttribute( RPT_ATTR_SUPPORT_1_OF_N );
   self.AddItem2ListView( 'Supports 1 of N', sValue, RPT_ATTR_SUPPORT_1_OF_N, EO_LOGIC, 00 );

end;

{
   This method retrieves section related information
   from the report and loads it into the grid.  It
   also associates the editing options available for
   each attribute.
}
procedure TfrmReport.LoadSectionInfo( pData : PTreeItemData );
var
   nSection : integer; // Current section of the report
   sValue   : string;  // String attribute data from the report
   nValue   : integer; // Numeric attribute data from the report

begin

   // Get the current section of the report.
   nSection := pData^.Section;

   // Get the attribute information from the report and fill in the listview.
   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_PAPER_SIZE );
   self.AddItem2ListView( 'Paper Size (DMPAPER_XXX Constant)', IntToStr( nValue ), SECTION_ATTR_PAPER_SIZE, EO_NEVER, 00 );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_PAPER_WIDTH );
   self.AddItem2ListView( 'Paper Width (TWIPS)', IntToStr( nValue ), SECTION_ATTR_PAPER_WIDTH, EO_NEVER, 00 );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_PAPER_LENGTH );
   self.AddItem2ListView( 'Paper Length (TWIPS)', IntToStr( nValue ), SECTION_ATTR_PAPER_LENGTH, EO_NEVER, 00 );

   sValue := self.xRpRun.GetSectionStringAttribute( nSection, SECTION_ATTR_LANDSCAPE );
   self.AddItem2ListView( 'Landscape', sValue, SECTION_ATTR_PAPER_LENGTH, EO_NEVER, 00 );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_PAPER_BIN );
   self.AddItem2ListView( 'Paper Bin (DMBIN_XXX Constant)', IntToStr( nValue ), SECTION_ATTR_PAPER_BIN, EO_NEVER, 00 );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_LEFT_MARGIN );
   self.AddItem2ListView( 'Left Margin (TWIPS)', IntToStr( nValue ), SECTION_ATTR_LEFT_MARGIN, EO_INPLACEEDIT, '9999' );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_TOP_MARGIN );
   self.AddItem2ListView( 'Top Margin (TWIPS)', IntToStr( nValue ), SECTION_ATTR_TOP_MARGIN, EO_INPLACEEDIT, '9999' );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_RIGHT_MARGIN );
   self.AddItem2ListView( 'Right Margin (TWIPS)', IntToStr( nValue ), SECTION_ATTR_RIGHT_MARGIN, EO_INPLACEEDIT, '9999' );

   nValue := self.xRpRun.GetSectionIntAttribute( nSection, SECTION_ATTR_BOTTOM_MARGIN );
   self.AddItem2ListView( 'Bottom Margin (TWIPS)', IntToStr( nValue ), SECTION_ATTR_BOTTOM_MARGIN, EO_INPLACEEDIT, '9999' );

   sValue := self.xRpRun.GetSectionStringAttribute( nSection, SECTION_ATTR_FILTER_EXP );
   self.AddItem2ListView( 'Filter', sValue, SECTION_ATTR_FILTER_EXP, EO_EXPBUILDER, EB_ENFORCE_LOGIC );

   sValue := self.xRpRun.GetSectionStringAttribute( nSection, SECTION_ATTR_SORT_ORDER_TEXT );
   self.AddItem2ListView( 'Sort Order', sValue, SECTION_ATTR_SORT_ORDER_TEXT, EO_INPLACEEDIT, 00 );

   sValue := self.xRpRun.GetSectionStringAttribute( nSection, SECTION_ATTR_SORT_ORDER_UNIQUE );
   self.AddItem2ListView( 'Unique Sort Order', sValue, SECTION_ATTR_SORT_ORDER_UNIQUE, EO_LOGIC, 00 );

end;

{
   This method retrieves table related information
   from the report and loads it into the grid.  It
   also associates the editing options available
   for each attribute.
}
procedure TfrmReport.LoadTableInfo( pData : PTreeItemData );
var
   nSection : integer; // Current section of the report
   sValue   : string;  // String attribute data from the report
   sTable   : string;  // Name of the table

begin

   // Get the information we saved in our structure.
   nSection := pData^.Section;
   sTable   := pData^.Name;

   // Each type of table has different properties depending on its technology.
   // We use the ReportPro class name to determine which attributes to retrieve.
   sValue := Trim( self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_CLASSNAME ) );

   // If this is a table accessed using an RDD (i.e. DBF's)...
   if ( sValue = 'rpRDDTable' ) then begin

      // Get the attribute information from the report and fill in the listview.
      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_DRIVER );
      self.AddItem2ListView( 'RDD', sValue, TABLE_ATTR_DRIVER, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_TABLE );
      self.AddItem2ListView( 'DBF File', sValue, TABLE_ATTR_TABLE, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_INDEX_FILE );
      self.AddItem2ListView( 'Index File', sValue, TABLE_ATTR_INDEX_FILE, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_INDEX_TAG );
      self.AddItem2ListView( 'Index Tag', sValue, TABLE_ATTR_INDEX_TAG, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_SEEK_EXPRESSION );
      self.AddItem2ListView( 'Seek Expression', sValue, TABLE_ATTR_SEEK_EXPRESSION, EO_EXPBUILDER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_WHILE_EXPRESSION );
      self.AddItem2ListView( 'While Expression', sValue, TABLE_ATTR_WHILE_EXPRESSION, EO_EXPBUILDER, EB_ENFORCE_LOGIC );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, TABLE_ATTR_FILTER_EXPRESSION );
      self.AddItem2ListView( 'Table Filter', sValue, TABLE_ATTR_FILTER_EXPRESSION, EO_EXPBUILDER, EB_ENFORCE_LOGIC );

   end

   // If this is a table accessed using an SQL query...
   else if ( sValue = 'rpSQLQuery' ) then begin

      // Get the attribute information from the report and fill in the listview.
      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_ODBC_SOURCE );
      self.AddItem2ListView( 'ODBC Source', sValue, SQLQUERY_ATTR_ODBC_SOURCE, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_NULL_AS_DEFAULT );
      self.AddItem2ListView( 'Null As Default', sValue, SQLQUERY_ATTR_SQL_NULL_AS_DEFAULT, EO_LOGIC, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_USER_COLS );
      self.AddItem2ListView( 'User Defined Columns', sValue, SQLQUERY_ATTR_SQL_USER_COLS, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_DISTINCT );
      self.AddItem2ListView( 'Distinct', sValue, SQLQUERY_ATTR_SQL_DISTINCT, EO_LOGIC, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_FROM );
      self.AddItem2ListView( 'From', sValue, SQLQUERY_ATTR_SQL_FROM, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_TABLE_WHERE );
      self.AddItem2ListView( 'Where (relational portion)', sValue, SQLQUERY_ATTR_SQL_TABLE_WHERE, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_FILTER_WHERE );
      self.AddItem2ListView( 'Where (filter portion)', sValue, SQLQUERY_ATTR_SQL_FILTER_WHERE, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_GROUP_BY );
      self.AddItem2ListView( 'Group by', sValue, SQLQUERY_ATTR_SQL_GROUP_BY, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_HAVING );
      self.AddItem2ListView( 'Having', sValue, SQLQUERY_ATTR_SQL_HAVING, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_UNION );
      self.AddItem2ListView( 'Union', sValue, SQLQUERY_ATTR_SQL_UNION, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_ORDERBY );
      self.AddItem2ListView( 'Order by', sValue, SQLQUERY_ATTR_SQL_ORDERBY, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLQUERY_ATTR_SQL_COL_DELIM );
      self.AddItem2ListView( 'Delimiter', sValue, SQLQUERY_ATTR_SQL_COL_DELIM, EO_NEVER, 00 );

   end

   // If this is a table accessed using SQL...
   else if ( sValue = 'rpSQLTable' ) then begin

      // Get the attribute information from the report and fill in the listview.
      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, SQLTABLE_ATTR_TABLE );
      self.AddItem2ListView( 'Table', sValue, SQLTABLE_ATTR_TABLE, EO_NEVER, 00 );

   end

   // If this is a table accessed using a Jasmine query...
   else if ( sValue = 'rpJasQuery' ) then begin

      // Get the attribute information from the report and fill in the listview.
      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, JASQUERY_ATTR_DATABASE );
      self.AddItem2ListView( 'Database', sValue, JASQUERY_ATTR_DATABASE, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, JASQUERY_ATTR_ENV_FILE );
      self.AddItem2ListView( 'Environment File', sValue, JASQUERY_ATTR_ENV_FILE, EO_NEVER, 00 );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, JASQUERY_ATTR_WHERE );
      self.AddItem2ListView( 'ODQL Where', sValue, JASQUERY_ATTR_WHERE, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, JASQUERY_ATTR_PREQUERY_ODQL );
      self.AddItem2ListView( 'ODQL Pre-query statements', sValue, JASQUERY_ATTR_PREQUERY_ODQL, EO_POPUPEDIT, false );

      sValue := self.xRpRun.GetTableStringAttribute( nSection, sTable, JASQUERY_ATTR_POSTQUERY_ODQL );
      self.AddItem2ListView( 'ODQL Post-query statements', sValue, JASQUERY_ATTR_POSTQUERY_ODQL, EO_POPUPEDIT, false );

   end;

end;

{
   This method retrieves variable related information from the
   report and loads it into the grid.  It also associates
   the editing options available for each attribute.
}
procedure TfrmReport.LoadVariableInfo( pData : PTreeItemData );
var
   nSection : integer; // Current section of the report
   sValue   : string;  // String attribute data from the report
   sName    : string;  // Name of the variable

begin

   // Get the information we saved in our structure.
   nSection := pData^.Section;
   sName    := pData^.Name;

   sValue := self.xRpRun.GetVariableStringAttribute(nSection, sName, VARIABLE_ATTR_RESET_LEVEL );
   self.AddItem2ListView( 'Reset At', sValue, VARIABLE_ATTR_RESET_LEVEL, EO_NEVER, 00 );

   // Get the attribute information from the report and fill in the listview.
   sValue := self.xRpRun.GetVariableStringAttribute( nSection, sName, VARIABLE_ATTR_INIT_EXPRESSION );
   self.AddItem2ListView( 'Initialization Expression', sValue, VARIABLE_ATTR_INIT_EXPRESSION, EO_EXPBUILDER, 00 );

   sValue := self.xRpRun.GetVariableStringAttribute(nSection, sName, VARIABLE_ATTR_UPDATE_LEVEL );
   self.AddItem2ListView( 'Update At', sValue, VARIABLE_ATTR_UPDATE_LEVEL, EO_NEVER, 00 );

   sValue := self.xRpRun.GetVariableStringAttribute(nSection, sName, VARIABLE_ATTR_UPDATE_EXPRESSION );
   self.AddItem2ListView( 'Update Expression', sValue, VARIABLE_ATTR_UPDATE_EXPRESSION, EO_EXPBUILDER, 00 );

end;

{******************************************************************************}
{****************************** Control Methods *******************************}
{******************************************************************************}

procedure TfrmReport.ClearEdits;
begin

   // Release the data stored in the treeview and listview items.
   self.ReleaseListViewData;
   self.ReleaseTreeViewData;

   // Clear the data out of the edit, treeview, and listview controls.
   self.edRpName.Clear;
   self.lvRpRun.Items.Clear;
   self.tvRpRun.Items.Clear;

end;

procedure TfrmReport.EnableControls;
begin

   // This method enables or disables the controls
   // depending on whether or not a report is loaded.

   // If there is nor report file name in the edit control...
   if ( self.edRpName.Text = '' ) then begin

      // Disable the controls.
      self.tvRpRun.Enabled       := false;
      self.lvRpRun.Enabled       := false;
      self.pbPrintDialog.Enabled := false;
      self.pbSetupDialog.Enabled := false;
      self.pbPreview.Enabled     := false;
      self.pbPrint.Enabled       := false;
      self.pbExport.Enabled      := false;

   end

   else begin

      // Enable the controls.
      self.tvRpRun.Enabled       := true;
      self.lvRpRun.Enabled       := true;
      self.pbPrintDialog.Enabled := true;
      self.pbSetupDialog.Enabled := true;
      self.pbPreview.Enabled     := true;
      self.pbPrint.Enabled       := true;
      self.pbExport.Enabled      := true;

   end;

end;

procedure TfrmReport.ReleaseTreeViewData;
var
   pData  : PTreeItemData; // Pointer to the data in the selected treeview item
   nItems : integer;       // Number of items in the treeview
   nItem  : integer;       // Loop counter

begin

   // Get the number of items in the treeview.
   nItems := self.tvRpRun.Items.Count - 01;

   // For each item in the treeview...
   for nItem := 00 to nItems do begin

      // Get a reference to the data in the item.
      pData := self.tvRpRun.Items[ nItem ].Data;

      // Release the memory for the stored structure.
      Dispose( pData );

   end;

end;

procedure TfrmReport.ReleaseListViewData;
var
   pData  : PListItemData; // Pointer to the data in the selected listview item
   nItems : integer;       // Number of items in the listview
   nItem  : integer;       // Loop counter

begin

   // Get the number of items in the listview.
   nItems := self.lvRpRun.Items.Count - 01;

   // For each item in the listview...
   for nItem := 00 to nItems do begin

      // Get a reference to the data in the item.
      pData := self.lvRpRun.Items[ nItem ].Data;

      // Release the memory for the stored structure.
      Dispose( pData );

   end;

end;

{******************************************************************************}
{***************************** Attribute Methods ******************************}
{******************************************************************************}

{
   This method is called to set a integer attribute in the report.
}
procedure TfrmReport.SetIntAttribute( nItemID, nSection : integer; sName : string; nAttribute, nValue : integer );
begin

   // Determine which method to call.
   case ( nItemID ) of
      RIT_REPORT,
      RIT_PRINTER  : self.xRpRun.SetReportIntAttribute( nAttribute, nValue );
      RIT_SECTION  : self.xRpRun.SetSectionIntAttribute( nSection, nAttribute, nValue );
      RIT_TABLE    : self.xRpRun.SetTableIntAttribute( nSection, sName, nAttribute, nValue );
      RIT_VARIABLE : ;
   end;

end;

{
   This method is called to set a string attribute in the report.
}
procedure TfrmReport.SetStringAttribute( nItemID, nSection : integer; sName : string; nAttribute : integer; sValue : string );
begin

   // Determine which method to call.
   case ( nItemID ) of
      RIT_REPORT,
      RIT_PRINTER  : self.xRpRun.SetReportStringAttribute( nAttribute, sValue );
      RIT_SECTION  : self.xRpRun.SetSectionStringAttribute( nSection, nAttribute, sValue );
      RIT_TABLE    : self.xRpRun.SetTableStringAttribute( nSection, sName, nAttribute, sValue );
      RIT_VARIABLE : self.xRpRun.SetVariableStringAttribute( nSection, sName, nAttribute, sValue );
   end;

end;

end.
