function varargout = form(varargin)
% FORM MATLAB code for form.fig
%      FORM, by itself, creates a new FORM or raises the existing
%      singleton*.
%      H = FORM returns the handle to a new FORM or the handle to
%      the existing singleton*.
%      FORM('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in FORM.M with the given input arguments.
%      FORM('Property','Value',...) creates a new FORM or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before form_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to form_OpeningFcn via varargin.
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
% See also: GUIDE, GUIDATA, GUIHANDLES
% Edit the above text to modify the response to help form
% Last Modified by GUIDE v2.5 11-Apr-2015 21:32:13
% Begin initialization code - DO NOT EDIT

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @form_OpeningFcn, ...
                   'gui_OutputFcn',  @form_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT

% --- Executes just before form is made visible.
function form_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to form (see VARARGIN)
% Choose default command line output for form
clc;


handles.output = hObject;
% Update handles structure
guidata(hObject, handles);

%handles buttons and textes at form open
set(handles.gsm1900,'Enable','off');
set(handles.gsm1900down,'Enable','off');
set(handles.text1,'String','No Instrument Yet');
set(handles.disconnect,'Enable','off');
set(handles.gsm900,'Enable','off');
set(handles.egsm900,'Enable','off');
set(handles.gsm900down,'Enable','off');
set(handles.gsm1800,'Enable','off');
set(handles.egsm900down,'Enable','off');
set(handles.gsm1800down,'Enable','off');
set(handles.usersettings,'Enable','off');
set(handles.text2,'ForegroundColor','Red');
set(handles.reset,'Enable','off');
set(handles.connect,'String','Connect & Capture');
set(handles.portnum,'String','Enter Port');
searchcomport=scanComPorts
set(handles.portfound,'String',searchcomport);

% --- Outputs from this function are returned to the command line.
function varargout = form_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% uiwait(msgbox('First choose the measurement option and then Connect & Trace'));
% Get default command line output from handles structure

varargout{1} = handles.output;

%warning about the need of antenna port number and sa ip address
%communication



% --- Executes on button press in connect.
function connect_Callback(hObject, eventdata, handles)
% hObject    handle to connect (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%handles buttons and textes after connecting to the analyzer
set(handles.gsm900,'Enable','on');
set(handles.egsm900,'Enable','on');
set(handles.gsm900down,'Enable','on');
set(handles.gsm1800,'Enable','on');
set(handles.egsm900down,'Enable','on');
set(handles.gsm1800down,'Enable','on');
set(handles.usersettings,'Enable','on');
set(handles.disconnect,'Enable','on');
set(handles.text2,'String','Connected To:')
set(handles.text2,'ForegroundColor','Green');
set(handles.reset,'Enable','on');
set(handles.connect,'String','Capture');
set(handles.gsm1900,'Enable','on');
set(handles.gsm1900down,'Enable','on');


%global variables for uses in many functions
global ip
global visaObj
global Attenuation Reference_Level Start_Frequency Stop_Frequency Resolution_BW Video_BW 
global Sweep_Number_Of_Points Sweep_Time Detector_Function Trace_Mode Scale_Type 
global Number_of_Averages Instrument_Model Instrument_Serial_Number Trace_data

%delete any instrument fidn
delete(instrfind)


%getting ip address and direction X,Y,Z for measurement
ip=get(handles.textip,'String')
antennadirection =get(handles.antennadir,'string');

%getting spectrum analyzer ip address from gui text
ip = strcat('TCPIP0::',ip,'::inst0::INSTR');
visaObj = visa('agilent',ip);    % for FSH8

%big buffer size for the png save image
visaObj.InputBufferSize = 100000;
% Set the timeout value
visaObj.Timeout = 10;
% Set the Byte order (not needed)
visaObj.ByteOrder = 'littleEndian';
%Open the Object
get(visaObj) %print properties

%opens the virtual instr obj
fopen(visaObj);

%%Instrument_Model and serial number first instruction
Instrument_string=query(visaObj,'*IDN?');
[Manufacturer,remain]=strtok(Instrument_string,',');
[Instrument_Model,remain]=strtok(remain,',');
[Serial_number,remain]=strtok(remain,',');
[Firmware_Version]=strtok(remain,',');
model=Instrument_Model

%putting the instrument string at textbox
set(handles.text1,'string', Instrument_string)

Instrument_Serial_Number=Serial_number;

error_exists=0;instrumentError='';

%finds and paste at textbox current measurement date
date=query(visaObj,':SYSTem:DATE?');
set(handles.date,'String',date);

fprintf(visaObj,'*CLS');
fprintf(visaObj,'ABORt');

%turn average off
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,'SWE:COUN 1');
else
    %Gia ton E4407B
    fprintf(visaObj,':SENSe:AVERage:STATe OFF')
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?'); %operation complete
fscanf(visaObj);

%set FSH8 in Spectrum analyzer mode if it not in this mode already
fprintf(visaObj,'INST:NSEL 1')
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
%pause(4)

%return the analyzer in clear Write Trace Mode
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,'DISP:WIND:TRAC:MODE WRIT');
else
    %Gia ton E4407B
    %fprintf(visaObj,[':sense:average:state OFF'])
    fprintf(visaObj,':TRAC:MODE WRITe');
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);   
%
fprintf(visaObj,':sense:detector:function POSitive');% default setting
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);

%Synchronous trace acquisition
fprintf(visaObj,'INITiate:IMMediate');
fprintf(visaObj,'INIT:CONT ON');
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);

%Clear any marker
fprintf(visaObj,':calculate:marker:STATe OFF');
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);

%pause(0.5)
%
%Frequency table creation
Start_Frequency=0; %str2num (query(visaObj,':sense:frequency:start?'));
Stop_Frequency=1000; %str2num (query(visaObj,':sense:frequency:stop?'));
Sweep_Number_Of_Points=631;

fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);

% %%ANTENNA POSITION
% %change the antenna direction before get measurement and after the set up
% %of analyzer
global searchcomport
antennadirection=get(handles.antennadir,'String');
antennaPos = antennadirection{get(handles.antennadir,'Value')};

% contents = get(handles.portnum,'String'); 
% antennadirection = contents{get(handles.portnum,'Value')};
% get(handles.portnum,'String')
% Rot(antennaPos,searchcomport);
Rot(antennaPos,get(handles.portnum,'String'))
%%


%% Taking Spectrum Analyzer current settings and paste them at review table
% 
%taking rf attenuation for review table
fprintf(visaObj,'input:ATTenuation?')
Char_Attenuation_setting=fscanf(visaObj,'input:ATTenuation?')
Atten=str2double(Char_Attenuation_setting)
set(handles.attenuationreview,'String',Atten)
set(handles.attenuationreview,'ForegroundColor','Green');

%taking refercelevel for review table
fprintf(visaObj,':DISPlay:WINDow:TRACe:Y:SCALe:RLEVel?')
Char_RL_setting=fscanf(visaObj,':DISPlay:WINDow:TRACe:Y:SCALe:RLEVel?')
Ref_Level=str2double(Char_RL_setting)
set(handles.reflevelreview,'String',Ref_Level)
set(handles.reflevelreview,'ForegroundColor','Green');


%taking resolution bandwidth for review table
fprintf(visaObj,':SENSe:BANDwidth:RESolution?')
Char_Res_BW=fscanf(visaObj,':SENSe:BANDwidth:RESolution?')
Res_BW=str2double(Char_Res_BW)/10^3 %in KHz
set(handles.resbandreview,'String',Res_BW)
set(handles.resbandreview,'ForegroundColor','Green');


%taking video bandwidth for review table
fprintf(visaObj,':SENSe:BANDwidth:VIDeo?')
Char_Video_Bandwidth=fscanf(visaObj,':SENSe:BANDwidth:VIDeo?')
Video_BW=str2double(Char_Video_Bandwidth)/10^6 %in MHz
set(handles.videobandreview,'String',Video_BW)
set(handles.videobandreview,'ForegroundColor','Green');

%taking start frequency for review table
fprintf(visaObj,':SENSe:FREQuency:STARt?')
Char_Start_Freq_setting=fscanf(visaObj,':SENSe:FREQuency:STARt?')
Start_Freq=str2double(Char_Start_Freq_setting)/10^6 %in MHz
set(handles.startfreqreview,'String',Start_Freq)
set(handles.startfreqreview,'ForegroundColor','Green');

%taking stop frequency for review table
fprintf(visaObj,':SENSe:FREQuency:STOP?')
Char_Stop_Freq_setting=fscanf(visaObj,':SENSe:FREQuency:STOP?')
Stop_Freq=str2double(Char_Stop_Freq_setting)/10^6 %in MHz
set(handles.stopfreqreview,'String',Stop_Freq)
set(handles.stopfreqreview,'ForegroundColor','Green');

%taking span frequency for review table
fprintf(visaObj,':SENSe:FREQuency:SPAN?')
Char_Span_Frequency=fscanf(visaObj,':SENSe:FREQuency:SPAN?')
Freq_Span=str2double(Char_Span_Frequency)/10^6 %in MHz
set(handles.spanfreqreview,'String',Freq_Span)
set(handles.spanfreqreview,'ForegroundColor','Green');

%taking sweep points for review table
Sweep_Points=631;
set(handles.sweepointsreview,'String',Sweep_Points)
set(handles.sweepointsreview,'ForegroundColor','Green');

%taking sweep time for review table
fprintf(visaObj,':SENSe:SWEep:TIME?')
Char_Sweep_Time=fscanf(visaObj,':SENSe:SWEep:TIME?')
Sweep_Time=str2double(Char_Sweep_Time)
set(handles.sweeptimereview,'String',num2str(Sweep_Time))
set(handles.sweeptimereview,'ForegroundColor','Green');

%taking detector function for review table
fprintf(visaObj,':sense:detector:function?')
Detector_Type=fscanf(visaObj,':sense:detector:function?')
set(handles.detectreview,'String',Detector_Type)
set(handles.detectreview,'ForegroundColor','Green');

%taking trace mode for review table
fprintf(visaObj,'DISP:WIND:TRAC:MODE?')
tracemode=fscanf(visaObj,'DISP:WIND:TRAC:MODE?')
set(handles.tracemodereview,'String',tracemode)
set(handles.tracemodereview,'ForegroundColor','Green');

%taking scale for review table
fprintf(visaObj,':DISPlay:WINDow:TRACe:Y:SPACing?')
scale=fscanf(visaObj,':DISPlay:WINDow:TRACe:Y:SPACing?')
set(handles.scalereview,'String',scale)
set(handles.scalereview,'ForegroundColor','Green');


%calling set measurement function for setting up the sa
[Start_Frequency,Stop_Frequency,Sweep_Number_Of_Points]=Set_measurement(visaObj,Atten,Ref_Level,Start_Freq,Stop_Freq,Res_BW,Video_BW,631,Sweep_Time,Detector_Type,tracemode,scale,5,Instrument_Model);


%%
%Get Trace data
Trace_data=[]
Trace_data=Get_trace_data(visaObj,Instrument_Model)

% fprintf(visaObj,'*WAI');
% fprintf(visaObj,'*OPC?');
% fscanf(visaObj)
    
%plot Trace
     plot_SA_Trace(visaObj,Trace_data);

% %Saving in excel the png plus measurement 
%frequency table creation
Freq_Step=(Stop_Frequency-Start_Frequency)/(Sweep_Number_Of_Points-1);
Freq_Table=[Start_Frequency:Freq_Step:Stop_Frequency]';
%%
%Put Measurement and analyzer screenshot in Excel file
 contents = get(handles.antennadir,'String'); 
 antennadirection = contents{get(handles.antennadir,'Value')};

Antena_Position=antennadirection
Antena_Kind='PCD 8250';
Cable_Kind='ARC both cables';
if strcmp(get(handles.startfreqreview,'String'),'880') || strcmp(get(handles.startfreqreview,'String'),'925')
    filename = strcat(Antena_Position,'_GSM900.xlsx');
else
    filename = strcat(Antena_Position,'_DCS1800.xlsx');
end

pathname=pwd;
sPut2Excel(visaObj,Instrument_Model,Instrument_Serial_Number,Trace_data,Freq_Table,Sweep_Number_Of_Points,Antena_Position,Antena_Kind,Cable_Kind,filename,pathname);

%%
%return the analyzer in clear Write Trace Mode
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    %fprintf(visaObj,[':SENSe:SWEep:COUNt 1'])
    fprintf(visaObj,'DISP:WIND:TRAC:MODE WRIT');
else
    %Gia ton E4407B
    fprintf(visaObj,[':sense:average:state OFF'])
    fprintf(visaObj,':TRAC:MODE WRITe')
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);   
%%
fprintf(visaObj,':sense:detector:function POSitive')% default setting
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
%%

% --- Executes on button press in disconnect.
function disconnect_Callback(hObject, eventdata, handles)
% hObject    handle to disconnect (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% hObject    handle to disconnect (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after disconnect
set(handles.text1,'String','')
set(handles.attenuationreview,'String','')
set(handles.reflevelreview,'String','')
set(handles.resbandreview,'String','')
set(handles.videobandreview,'String','')
set(handles.startfreqreview,'String','')
set(handles.stopfreqreview,'String','')
set(handles.spanfreqreview,'String','')
set(handles.sweepointsreview,'String','')
set(handles.date,'String','');
set(handles.sweeptimereview,'String','')
set(handles.detectreview,'String','')
set(handles.scalereview,'String','')
set(handles.tracemodereview,'String','')
set(handles.text2,'ForegroundColor','Red');
set(handles.text2,'String','Disconnected');
set(handles.gsm900,'Enable','off');
set(handles.egsm900,'Enable','off');
set(handles.gsm900down,'Enable','off');
set(handles.gsm1800,'Enable','off');
set(handles.egsm900down,'Enable','off');
set(handles.gsm1800down,'Enable','off');
set(handles.usersettings,'Enable','off');
set(handles.connect,'Enable','on');
set(handles.disconnect,'Enable','off');
set(handles.text1,'String','Disconnected');
set(handles.reset,'Enable','off');
set(handles.connect,'ForegroundColor','BLACK');
set(handles.connect,'String','Connect & Capture');
set(handles.portnum,'String','');
set(handles.gsm1900,'Enable','off');
set(handles.gsm1900down,'Enable','off');

global visaObj ip
if isempty(visaObj)
    ip = strcat('TCPIP0::',ip,'::inst0::INSTR');
    visaObj = visa('agilent',ip); 
else
    fclose(visaObj);
    visaObj = visaObj(1);
end

fclose(visaObj);
clear all;
%%

% --- Executes on button press in Exit.
function close_Callback(hObject, eventdata, handles)
% hObject    handle to Exit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

global visaObj;
global ip;
if isempty(visaObj)
%     saip = strcat('TCPIP0::',saip,'::inst0::INSTR');
    visaObj = visa('agilent',ip); % for FSH8
else
    fclose(visaObj);
    visaObj = visaObj(1);
end

clc;
fclose(visaObj);
clear all;
CLOSE ALL;
% Clean up all objects.
delete(visaObj);
close form;
%%

% --------------------------------------------------------------------
function Exit_Callback(hObject, eventdata, handles)
% hObject    handle to Exit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% visaObj = instrfind('agilent','TCPIP1');

clc;
clear all;
close all;
%%

% --- Executes on button press in reset.
function reset_Callback(hObject, eventdata, handles)
% hObject    handle to reset (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

global visaObj;
global ip;
if isempty(visaObj)
    %saip = strcat('TCPIP0::',saip,'::inst0::INSTR');
    visaObj = visa('agilent',ip); % for FSH8
else
    fclose(visaObj);
    visaObj = visaObj(1);
end

fopen(visaObj);
fprintf(visaObj,'SYST:PRES:FACT');
fclose(visaObj);
fprintf(visaObj,'*RST');
%%

% --- Executes on button press in evaluate.
function evaluate_Callback(hObject, eventdata, handles)
% hObject    handle to evaluate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%calls calculate e function for evaluate measurement and results
Calculate_E;
%%

% --------------------------------------------------------------------
function usersettings_Callback(hObject, eventdata, handles)
% hObject    handle to usersettings (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%uiwait for warning about the correct set up for measure
uiwait(warndlg(sprintf('Warning !! \nFor Detection Function you can use AVERage | POSitive | QUASipeak | RMS \nFor Scale Type you can use LOG or LIN\nFor Trace Mode you can use AVERage | MAXHold | MINHold | VIEW | WRITe')));

prompt={'Enter Start Frequency (MHz)','Enter Stop Frequency (MHz)'...
    ,'Enter Attenuation (dB)','Enter Reference Level (dBm)'...
    ,'Enter Resolution Bandwidth (KHz)','Enter Video Bandwidth (MHz)'...
    ,'Enter Sweep Time (Seconds)','Enter Sweep Points'...
    ,'Enter Span Frequency (MHz)','Enter Detector Function'...
    ,'Enter Scale Type (Log,Lin)','Enter Trace Mode'};
name='Costum Measurement';
numlines=1;
defaultanswer={'0','0','0','0','0','0','0','0','0','RMS','LOG','AVER'};
answer=inputdlg(prompt,name,numlines,defaultanswer);

%Set up gui after costum measurement
set(handles.startfreqreview,'String',answer{1});
set(handles.stopfreqreview,'String',answer{2});
set(handles.attenuationreview,'String',answer{3});
set(handles.reflevelreview,'String',answer{4});
set(handles.resbandreview,'String',answer{5});
set(handles.videobandreview,'String',answer{6});
set(handles.sweeptimereview,'String',answer{7});
set(handles.sweepointsreview,'String',answer{8});
set(handles.spanfreqreview,'String',answer{9});
set(handles.detectreview,'String',answer{10});
set(handles.scalereview,'String',answer{11});
set(handles.tracemodereview,'String',answer{12});
set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,answer{3},answer{4},answer{1},answer{2},answer{5},answer{6},answer{8},answer{7},answer{10},answer{12},answer{11},100,Instrument_Model);
%%

% --- Executes on button press in gsm900.
function gsm900_Callback(hObject, eventdata, handles)
% hObject    handle to gsm900 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','890');
% set(handles.stopfreqreview,'String','960');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','90');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,890,960,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model);
%%

% --- Executes on button press in egsm900.
function egsm900_Callback(hObject, eventdata, handles)
% hObject    handle to egsm900 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','880');
% set(handles.stopfreqreview,'String','960');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','40');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,880,960,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model);

%%

% --- Executes on button press in gsm900down.
function gsm900down_Callback(hObject, eventdata, handles)
% hObject    handle to gsm900down (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','935');
% set(handles.stopfreqreview,'String','960');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','45');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,935,960,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model)

%%

% --- Executes on button press in gsm1800.
function gsm1800_Callback(hObject, eventdata, handles)
% hObject    handle to gsm1800 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','1710');
% set(handles.stopfreqreview,'String','1880');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','200');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,1710,1880,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model)

%%

% --- Executes on button press in egsm900down.
function egsm900down_Callback(hObject, eventdata, handles)
% hObject    handle to egsm900down (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','925');
% set(handles.stopfreqreview,'String','960');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','100');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,925,960,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model)
%%

% --- Executes on button press in gsm1800down.
function gsm1800down_Callback(hObject, eventdata, handles)
% hObject    handle to gsm1800down (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','1805');
% set(handles.stopfreqreview,'String','1880');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','100');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,1805,1880,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model)

%%


% --------------------------------------------------------------------
function About_Callback(hObject, eventdata, handles)
% hObject    handle to About (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
open about.fig;

function options_Callback(hObject, eventdata, handles)
% hObject    handle to options (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% --------------------------------------------------------------------

function moreabout_Callback(hObject, eventdata, handles)
% hObject    handle to moreabout (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
web('http://www.mathworks.com/matlabcentral/fileexchange/index?utf8=%E2%9C%93&term=gui+instrument+control','-browser');

% --- Executes on selection change in antennadir.
function antennadir_Callback(hObject, eventdata, handles)
% hObject    handle to antennadir (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hints: contents = cellstr(get(hObject,'String')) returns antennadir contents as cell array
%        contents{get(hObject,'Value')} returns selected item from antennadir
% --- Executes during object creation, after setting all properties.

function antennadir_CreateFcn(hObject, eventdata, handles)
% hObject    handle to antennadir (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes during object creation, after setting all properties.
function textip_CreateFcn(hObject, eventdata, handles)
% hObject    handle to textip (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

function portnum_Callback(hObject, eventdata, handles)
% hObject    handle to portnum (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hints: get(hObject,'String') returns contents of portnum as text
%        str2double(get(hObject,'String')) returns contents of portnum as a double
% --- Executes during object creation, after setting all properties.

function portnum_CreateFcn(hObject, eventdata, handles)
% hObject    handle to portnum (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function textip_Callback(hObject, eventdata, handles)
% hObject    handle to textip (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hints: get(hObject,'String') returns contents of textip as text
%        str2double(get(hObject,'String')) returns contents of textip as a double



function portfound_Callback(hObject, eventdata, handles)
% hObject    handle to portfound (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of portfound as text
%        str2double(get(hObject,'String')) returns contents of portfound as a double


% --- Executes during object creation, after setting all properties.
function portfound_CreateFcn(hObject, eventdata, handles)
% hObject    handle to portfound (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in gsm1900.
function gsm1900_Callback(hObject, eventdata, handles)
% hObject    handle to gsm1900 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% set(handles.startfreqreview,'String','1850');
% set(handles.stopfreqreview,'String','1990');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','100');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,1850,1990,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model)

% --- Executes on button press in gsm1900down.
function gsm1900down_Callback(hObject, eventdata, handles)
% hObject    handle to gsm1900down (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%Set up gui after auto measurement
% must be commends after the trace button update
% set(handles.startfreqreview,'String','1930');
% set(handles.stopfreqreview,'String','1990');
% set(handles.attenuationreview,'String','10');
% set(handles.reflevelreview,'String','0');
% set(handles.resbandreview,'String','100');
% set(handles.videobandreview,'String','1');
% set(handles.sweeptimereview,'String','100');
% set(handles.sweepointsreview,'String','631');
% set(handles.spanfreqreview,'String','100');
% set(handles.detectreview,'String','RMS');
% set(handles.scalereview,'String','LOG');
% set(handles.tracemodereview,'String','AVER');
 set(handles.connect,'ForegroundColor','Red');

%setting up sa for trace
global Instrument_Model visaObj
Set_measurement(visaObj,10,0,1930,1990,100000,1000000,631,0.1,'RMS','AVER','LOG',100,Instrument_Model)
