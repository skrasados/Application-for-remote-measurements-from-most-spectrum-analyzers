function [ output_args ] = sPut2Excel(visaObj,Instrument_Model,Instrument_Serial_Number,Trace_data,Freq_Table,Sweep_Number_Of_Points,Antena_Position,Antena_Kind,Cable_Kind,filename,pathname)
  
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    Attenuation=str2double(query(visaObj,':INPut:ATTenuation?'));
else
    %Gia ton E4407B
    Attenuation=str2double(query(visaObj,':SENSe:POWer:RF:ATTenuation?'));
end
%%
%%Center_Frequency
Center_Frequency=str2double(query(visaObj,':SENSe:FREQuency:CENTer?'));%in Hz
%%    
%%Date_Time
SA_Date=query(visaObj,':SYSTem:DATE?');%,'dd/mm/yyyy'))% HH:MM'))
SA_Date=datestr(SA_Date,'dd/mm/yyyy');
SA_Time=query(visaObj,':SYSTem:TIME?');%'HH:MM:SS'
[Time_in_hours,remain]=strtok(SA_Time, ',');
[Time_in_minutes,remain]=strtok(remain, ',');
Time_in_minutes=strtrim(Time_in_minutes);
SA_Time=[Time_in_hours,':',Time_in_minutes];
Date_Time=[SA_Date ' ' SA_Time];
%%
%%
%Reference_Level
Reference_Level=str2double(query(visaObj,':DISPlay:WINDow:TRACe:Y:SCALe:RLEVel?'));
%%   
%%Resolution_BW
Resolution_BW=str2double(query(visaObj,':SENSe:BANDwidth:RESolution?'));
%%    
%%Scale_Type den paizoyn aytes sto FSH8
%fprintf(visaObj,':DISPlay:WINDow:TRACe:Y:SCALe:SPACing?')
%Scale_Type=fscanf(visaObj,':DISPlay:WINDow:TRACe:Y:SCALe:SPACing?')

%sto FSH8 paizoyn aytes ayti i opoia paizei kai ston E4407B
%length(Scale_Type)% dinei tesseris characters
Scale_Type = strtrim(query(visaObj,':DISPlay:WINDow:TRACe:Y:SPACing?'));
%length(Scale_Type)%dinei treis characters
%%    
%%Span_Frequency
Span_Frequency=str2double(query(visaObj,':SENSe:FREQuency:SPAN?'));
%%
%%Start_Frequency
Start_Frequency=str2double(query(visaObj,':SENSe:FREQuency:STARt?'));
%%    
%%Stop_Frequency
Stop_Frequency=str2double(query(visaObj,':SENSe:FREQuency:STOP?'));
%%    
%%Sweep_Number_Of_Points  den paizei sto FSH8. Se ayto ;exoyme 631 sweep
%%points. paizei ayti
%fprintf(visaObj,'SWE:POIN')
%alla den apizei ayti
%fscanf(visaObj,'SWE:POIN?')
%fprintf(visaObj,':SENSe:SWEep:POINts?')
%Char_Sweep_Points=fscanf(visaObj,':SENSe:SWEep:POINts?')
%Sweep_Number_Of_Points=str2double(Char_Sweep_Points)
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    Sweep_Number_Of_Points=631;
else
    %Gia ton E4407B
    Sweep_Number_Of_Points=str2double(query(visaObj,':SENSe:SWEep:POINts?'));
end
%%    
%Sweep_Time
Sweep_Time=str2double(query(visaObj,':SENSe:SWEep:TIME?'));
%%    
%Video_BW
Video_BW=str2double(query(visaObj,':SENSe:BANDwidth:VIDeo?'));
%%
    if isequal(filename,0) || isequal(pathname,0)
       %msgbox('User pressed cancel')
       return
    else
       %msgbox(['User selected ', fullfile(pathname, filename)]);
    end

%%    
% First open an Excel Server
Excel = actxserver('Excel.Application');
set(Excel, 'Visible', 1);
%get(Excel);

% Insert a new workbook
Workbooks = Excel.Workbooks;
Workbook = invoke(Workbooks, 'Add');
%
% Make the first sheet active
Sheets = Excel.ActiveWorkBook.Sheets;
sheet1 = get(Sheets, 'Item', 1);
invoke(sheet1, 'Activate');
%
% Get a handle to the active sheet
Activesheet = Excel.Activesheet;
%onomazei to sheet tou excel se analyzer data
set(Activesheet,'name','Analyser Data');
%%
% Put a MATLAB array into Excel

%onomasia keliou frequency A1
ActivesheetRange = get(Activesheet,'Range','A1:A1');
set(ActivesheetRange, 'Value', 'Frequency (Hz)');
%onomasia keliou trace B1
ActivesheetRange = get(Activesheet,'Range','B1:B1');
set(ActivesheetRange, 'Value', 'Trace1 (dBm)');
%vazei sthn sthlh a katw apo thn onomasia keliou a1 olo ton pinaka
%Freq_Table
ActivesheetRange = get(Activesheet,'Range',['A2:A' num2str(Sweep_Number_Of_Points+1)]);
set(ActivesheetRange, 'Value', Freq_Table);
%vazei sthn sthlh a katw apo thn onomasia keliou b1 olo ton pinaka
%trace data
ActivesheetRange = get(Activesheet,'Range',['B2:B' num2str(Sweep_Number_Of_Points+1)]);
set(ActivesheetRange, 'Value', [Trace_data]);
%onomasia keliou attenuation C1
ActivesheetRange = get(Activesheet,'Range','C1:C1');
set(ActivesheetRange, 'Value', 'Attenuation (dB)');
%vazei sthn sthlh a katw apo thn onomasia keliou c1 olo ton pinaka
%me ta attenuation
ActivesheetRange = get(Activesheet,'Range','C2:C2');
set(ActivesheetRange, 'Value', Attenuation);
%onomasia keliou attenuation C4:C4
ActivesheetRange = get(Activesheet,'Range','C4:C4');
set(ActivesheetRange, 'Value', 'Center Frequency (Hz)');
%vazei sto keli katw apo to c4c4 to center frequency
ActivesheetRange = get(Activesheet,'Range','C5:C5');
set(ActivesheetRange, 'Value', Center_Frequency);
%onomasia keliou date C7:C7
ActivesheetRange = get(Activesheet,'Range','C7:C7');
set(ActivesheetRange, 'Value', 'Date/Time');
%vazei sto keli katw apo to c7c7 thn hmeromhnia
ActivesheetRange = get(Activesheet,'Range','C8:C8');
set(ActivesheetRange, 'Value', Date_Time);
%onomasia keliou C10C10 Instrument_Model
ActivesheetRange = get(Activesheet,'Range','C10:C10');
set(ActivesheetRange, 'Value', 'Instrument Model');
%vazei sto keli katw apo to c10c10 to modelo tou analyth
ActivesheetRange = get(Activesheet,'Range','C11:C11');
set(ActivesheetRange, 'Value',  Instrument_Model);
%onomasia keliou C13C13 Instrument_serial number
ActivesheetRange = get(Activesheet,'Range','C13:C13');
set(ActivesheetRange, 'Value', 'Instrument Serial Number');
%vazei sto keli katw apo to c13c13 to modelo tou analyth
ActivesheetRange = get(Activesheet,'Range','C14:C14');
set(ActivesheetRange, 'Value', Instrument_Serial_Number);

ActivesheetRange = get(Activesheet,'Range','C16:C16');
set(ActivesheetRange, 'Value', 'Reference Level (dBm)');
ActivesheetRange = get(Activesheet,'Range','C17:C17');
set(ActivesheetRange, 'Value', Reference_Level);

ActivesheetRange = get(Activesheet,'Range','C19:C19');
set(ActivesheetRange, 'Value', 'Resolution BW (Hz)');
ActivesheetRange = get(Activesheet,'Range','C20:C20');
set(ActivesheetRange, 'Value', Resolution_BW);

ActivesheetRange = get(Activesheet,'Range','C22:C22');
set(ActivesheetRange, 'Value', 'Scale Type');
ActivesheetRange = get(Activesheet,'Range','C23:C23');
set(ActivesheetRange, 'Value', Scale_Type);

ActivesheetRange = get(Activesheet,'Range','C25:C25');
set(ActivesheetRange, 'Value', 'Span Frequency (Hz)');
ActivesheetRange = get(Activesheet,'Range','C26:C26');
set(ActivesheetRange, 'Value', Span_Frequency);

ActivesheetRange = get(Activesheet,'Range','C28:C28');
set(ActivesheetRange, 'Value', 'Start Frequency (Hz)');
ActivesheetRange = get(Activesheet,'Range','C29:C29');
set(ActivesheetRange, 'Value', Start_Frequency);

ActivesheetRange = get(Activesheet,'Range','C31:C31');
set(ActivesheetRange, 'Value', 'Stop Frequency (Hz)');
ActivesheetRange = get(Activesheet,'Range','C32:C32');
set(ActivesheetRange, 'Value', Stop_Frequency);

ActivesheetRange = get(Activesheet,'Range','C34:C34');
set(ActivesheetRange, 'Value', 'Sweep Number Of Points');
ActivesheetRange = get(Activesheet,'Range','C35:C35');
set(ActivesheetRange, 'Value', Sweep_Number_Of_Points);

ActivesheetRange = get(Activesheet,'Range','C37:C37');
set(ActivesheetRange, 'Value', 'Sweep Time (seconds)');
ActivesheetRange = get(Activesheet,'Range','C38:C38');
set(ActivesheetRange, 'Value', Sweep_Time);

ActivesheetRange = get(Activesheet,'Range','C40:C40');
set(ActivesheetRange, 'Value', 'Video BW (Hz)');
ActivesheetRange = get(Activesheet,'Range','C41:C41');
set(ActivesheetRange, 'Value', Video_BW);
%%
ActivesheetRange = get(Activesheet,'Range','C55:C55');
set(ActivesheetRange, 'Value', 'Kind of Antenna');
ActivesheetRange = get(Activesheet,'Range','C56:C56');
set(ActivesheetRange, 'Value', 'PCD 8250');
ActivesheetRange = get(Activesheet,'Range','C58:C58');
set(ActivesheetRange, 'Value', 'Antenna Polarization');
ActivesheetRange = get(Activesheet,'Range','C59:C59');
set(ActivesheetRange, 'Value', Antena_Position);
%%
% Get a handle to the active sheet
%%orizei to upsos tou excel sheet na exei telos ekei pou teleionoun ta
%%sweep points +1 gia na fenetai omorfa
ActivesheetRange = get(Activesheet,'Range',['A1:C' num2str(Sweep_Number_Of_Points+1)]);
%******************set(ActivesheetRange, 'HorizontalAlignment', 'xlCenter');
%get(ActivesheetRange);
%************align_H=get(ActivesheetRange,'HorizontalAlignment')
%***********align_V=get(ActivesheetRange,'VerticalAlignment')
set(ActivesheetRange,'HorizontalAlignment',-4108);
set(ActivesheetRange,'VerticalAlignment',-4108);
set(ActivesheetRange,'ColumnWidth',22);
%%
% Make the second sheet active
%Sheets = Excel.ActiveWorkBook.Sheets;

%anoigei to neo sheet tou excel wste na mporesei na apothikeysh to png
%arxeio
sheet2 = get(Sheets, 'Item', 2);
invoke(sheet2, 'Activate');
%%
% Get a handle to the active sheet
%to onomazei ws analyzer screen to sheet
Activesheet = Excel.Activesheet;
set(Activesheet,'name','Analyser Screen');
%%
%Put the image in the second Sheet of Excel file
%kalei thn spic2xls gia na apothikeusei sto sheet thn png eikona
spic2xls(Excel,visaObj,Instrument_Model,'',strcat(pathname,filesep,filename),'Analyser Screen',[0 0 640 480]);
%%
%Now save the workbook
%invoke(Workbook, 'SaveAs', 'myfile.xls'); %to paei sto my documents
invoke(Workbook, 'SaveAs', fullfile(pathname, filename)); %kanei save kai to paei ekei poy epilexame
% To avoid saving the workbook and being prompted to do so,
% uncomment the following code.
 Workbook.Saved = 1;
%% 
 invoke(Workbook, 'Close');
%%
% Quit Excel
invoke(Excel, 'Quit');
%%
% End process
delete(Excel);    
msgbox(['The data were saved at ', fullfile(pathname, filename)]);

end

