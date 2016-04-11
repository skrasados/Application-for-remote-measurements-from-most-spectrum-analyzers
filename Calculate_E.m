function varargout = Calculate_E(varargin)
% CALCULATE_E MATLAB code for Calculate_E.fig
%      CALCULATE_E, by itself, creates a new CALCULATE_E or raises the existing
%      singleton*.
%
%      H = CALCULATE_E returns the handle to a new CALCULATE_E or the handle to
%      the existing singleton*.
%
%      CALCULATE_E('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in CALCULATE_E.M with the given input arguments.
%
%      CALCULATE_E('Property','Value',...) creates a new CALCULATE_E or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Calculate_E_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Calculate_E_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Calculate_E

% Last Modified by GUIDE v2.5 18-Apr-2015 14:16:35

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Calculate_E_OpeningFcn, ...
                   'gui_OutputFcn',  @Calculate_E_OutputFcn, ...
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


% --- Executes just before Calculate_E is made visible.
function Calculate_E_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Calculate_E (see VARARGIN)

% Choose default command line output for Calculate_E
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);
% percent = percentfor{get(handles.popupmenu2,'Value')};

% UIWAIT makes Calculate_E wait for user response (see UIRESUME)
% uiwait(handles.figure2);


% --- Outputs from this function are returned to the command line.
function varargout = Calculate_E_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in calculate.
function calculate_Callback(hObject, eventdata, handles)
% hObject    handle to calculate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)



% Read the the trace files for the three dimensions (X-Y-Z).
%Sth sthlh 1 mpainei to euros twn syxnothtwn(se MHz) sto opoio egine h metrhsh. Sth sthlh 2,3,4 tou 
%EXCELS_File mpainei h isxys se dB opws mas thn edwse o analyths gia tis 3
%diaforetikes kateuthinseis ths keraias

[name,PathName] = uigetfile('*.xlsx','Select the X-Spectrum Analyzer Trace file');
temp = xlsread(name,1);
EXCELS_File(:,1) = temp(:,1) ;
EXCELS_File(:,2) = temp(:,2);
[name,PathName] = uigetfile('*.xlsx','Select the Y-Spectrum Analyzer Trace file');
temp = xlsread(name,1);
EXCELS_File(:,3) = temp(:,2);
[name,PathName] = uigetfile('*.xlsx','Select the Z-Spectrum Analyzer Trace file');
temp = xlsread(name,1);
EXCELS_File(:,4) = temp(:,2);

start_freq=temp(28,3) * 10^(-6);
stop_freq=temp(31,3) * 10^(-6) ;
number_of_points= temp(34,3);
Interpolation_step=(stop_freq-start_freq)/(number_of_points-1);

resolutionbandwidth = temp(19,3) * 10^(-6);
% RBW(MHz) filtro 8oryvou tou analyth syntelestis k gia ton FSH8 einai 5 kai gia
% to Agilent 1.128
bandwidthnoise = 5 * resolutionbandwidth ; 


% B_ch(MHz) einai to integration Bandwidth :(stop_freq - start freq) gia to
bandwidthchannel = stop_freq - start_freq;

% Ypologismos tou syntelesth dior8wshs(se dB)
%Ypologismos correction factor
ChannelPw = 10 * log10( (bandwidthchannel/bandwidthnoise)* (1/number_of_points) ); 

% to f_ einai se MHz
f_i = EXCELS_File(:,1);

f_i = f_i/1000000; %SE HZ

[File_name,Path_name] = uigetfile('*.xlsx','Select the AntennaFactor and Cable losses file')
Isotropic_PCD_all_table = xlsread('ISO_SWR_UN_PCD.xls');

if (start_freq==880) %%gia gsm900
    temp_ = xlsread(File_name);
    Gain_Losses = temp_(801:891,:) ; %sthlh gia sixnothtes 880-970 mhz
    Isotropic_PCD_corrections_hz=Isotropic_PCD_all_table(801:891,:);

elseif (start_freq==925) %%gia gsm900dl
    temp_ = xlsread(File_name);
    Gain_Losses = temp_(846:891,:) ;
    Isotropic_PCD_corrections_hz=Isotropic_PCD_all_table(846:891,:);

elseif (start_freq==1710) %%gia gsm1800
     temp_ = xlsread(File_name);
     Gain_Losses = temp_(1631:1801,:) ;  %sthlh gia sixnothtes 1700-1900 mhz
     Isotropic_PCD_corrections_hz=Isotropic_PCD_all_table(1631:1801,:);

elseif (start_freq==1850) %%gia gsm1900
     temp_ = xlsread(File_name);
     Gain_Losses = temp_(1771:1911,:) ;
     Isotropic_PCD_corrections_hz=Isotropic_PCD_all_table(1771:1911,:);

elseif (start_freq==1930) %%gia gsm1900dl
     temp_ = xlsread(File_name);
     Gain_Losses = temp_(1851:1911,:) ; 
     Isotropic_PCD_corrections_hz=Isotropic_PCD_all_table(1851:1911,:);
else  %%gia gsm1800dl
    temp_ = xlsread(File_name);
    Gain_Losses = temp_(1726:1801,:) ;
    Isotropic_PCD_corrections_hz=Isotropic_PCD_all_table(1726:1801,:);

end

w = start_freq:Interpolation_step:stop_freq; % frequencies table in MHz

Isotropic_PCD_corrections_A = interp1(Isotropic_PCD_corrections_hz(:,2),Isotropic_PCD_corrections_hz(:,3),w,'linear');
Isotropic_PCD_corrections_B = interp1(Isotropic_PCD_corrections_hz(:,2),Isotropic_PCD_corrections_hz(:,4),w,'linear');

Isotropy=(Isotropic_PCD_corrections_A).* f_i' + Isotropic_PCD_corrections_B;
Isotropy_linear_antenna = 10 .^(Isotropy/20);



P_i_x = EXCELS_File(:,2); %measured X dBm
P_i_y = EXCELS_File(:,3); %measured Y dBm
P_i_z = EXCELS_File(:,4); %measured Z dBm

% Ypologismos Isxyos sthn eisodo tou syshmatos(se ka8e stoixeio tou p_out prosti8etai to p_c)
P_out_x = P_i_x + ChannelPw;
P_out_y = P_i_y + ChannelPw;
P_out_z = P_i_z + ChannelPw;


% Grammikh Paremvolh ap ton pinaka me ta AF 
x = start_freq:Interpolation_step:stop_freq; % frequencies table in MHz
AF_interpolated(:,1) = interp1(Gain_Losses(:,2),Gain_Losses(:,3),x,'linear')
AF_db = AF_interpolated(:,1);

AF_Absolute = (10 .^ (AF_db / 20)) .* Isotropy_linear_antenna';
%AF_dB(i) = 20 * Log(AF_Absolute(i)) / Log(10)
AF_db = 20 * log10(AF_Absolute);

Gain_db = -29.776613 + 20 * log10(f_i) - AF_db;

% Grammikh Paremvolh ap ton pinaka me ta losses
w = start_freq:Interpolation_step:stop_freq; % frequencies table in MHz
CableLosses_interpolated(:,1) = interp1(Gain_Losses(:,2),Gain_Losses(:,9),w,'linear');
Cablelosses = CableLosses_interpolated(:,1);

%euresh l total
L_total = Cablelosses - Gain_db; 


%Sin,dbm poiknotitas isxios stin eisodo X
S_in_xdBm = -48.76 + P_out_x + L_total + AF_db;
%Sin,dbm poiknotitas isxios stin eisodo Y
S_in_ydBm = -48.76 + P_out_y + L_total + AF_db;
%Sin,dbm poiknotitas isxios stin eisodo Z
S_in_zdBm = -48.76 + P_out_z + L_total + AF_db;


%S(W) = 1W · 10(P(dBm) / 10) / 1000 = 10((P(dBm) - 30) / 10)

% S_in_xWm = ( 10.^(S_in_xdBm/10) / 1000 );
% S_in_yWm = ( 10.^(S_in_ydBm/10) / 1000 );
% S_in_zWm = ( 10.^(S_in_zdBm/10) / 1000 );

S_in_xWm = 10 * 10.^(S_in_xdBm/10);
S_in_yWm = 10 * 10.^(S_in_ydBm/10);
S_in_zWm = 10 * 10.^(S_in_zdBm/10);

E_in_x = sqrt(120 * 3.1416 .* S_in_xWm); %E_in_x entasi hlektrikou gia X
E_in_y = sqrt(120 * 3.1416 .* S_in_yWm); %E_in_y entasi hlektrikou gia y
E_in_z = sqrt(120 * 3.1416 .* S_in_zWm); %E_in_z entasi hlektrikou gia z


H_in_x = sqrt(S_in_xWm ./ (120 * 3.1416)); %H_in_x entasi magnhtikou gia X
H_in_y = sqrt(S_in_yWm ./ (120 * 3.1416)); %H_in_x entasi magnhtikou gia Y
H_in_z = sqrt(S_in_zWm ./ (120 * 3.1416)); %H_in_x entasi magnhtikou gia Z


E_oliko_fi = sqrt(E_in_x + E_in_y + E_in_z);

E_total = sqrt(sum(E_oliko_fi.^2));

S_total = E_total^2 / (120 * 3.1416);

global limitation
global percent

%elegxos ti pedio sixnothtwn tha ginei evaluation me tis global times pou
%eginan set
if strcmp(limitation,'gsm')
    if strcmp(percent,'70')
        Limit = 34.5; %gsm900 me 70% tou ICNIRP        
    elseif strcmp(percent,'60')
        Limit = 31.9; %gsm900 me 60% tou ICNIRP
    elseif strcmp(percent,'100')
        Limit = 41.2; %gsm900 me 100% tou ICNIRP
    else
        warndlg('You have to set 60 for sensitive areas or 70 percent any normal area','!! Warning !!')
    end
elseif strcmp(limitation,'dcs')
    if strcmp(percent,'70')
        Limit = 48.8; %gsm1800 me 70% tou ICNIRP
    elseif strcmp(percent,'60')
        Limit = 45.2; %gsm1800 me 60% tou ICNIRP 
    elseif strcmp(percent,'100')
        Limit = 58.2; %gsm900 me 100% tou ICNIRP
    else
        warndlg('You have to set 60 for sensitive areas or 70 percent any normal area','!! Warning !!')
    end  
else
        warndlg('Set up please the correct limitation, gsm or dcs','!! Warning !!')

end


Logos_Hlektrikou_fi = (E_oliko_fi).^2 / (Limit).^2; %logos E_i/E_Limit
Syntelestis_ekthesis = sum(Logos_Hlektrikou_fi);


if (Syntelestis_ekthesis <= 1) %%?
    h = msgbox('Within Limitations','title'); %popup messagebox gia enhmerwsh xrhsth oso aforia an eimaste entos h ektos oriwn
else
    h = msgbox('Out Of Bounds','title','Warn');  %popup messagebox gia enhmerwsh xrhsth oso aforia an eimaste entos h ektos oriwn
end

file=strcat(limitation,'_results.xls');
filename = file;
pathname=pwd;

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
%
set(Activesheet,'name','Results');
%%
% Put a MATLAB array into Excel

ActivesheetRange = get(Activesheet,'Range','A1:A1');
set(ActivesheetRange, 'Value', 'Frequency(MHz)');
ActivesheetRange = get(Activesheet,'Range',['A2:A' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', f_i);

ActivesheetRange = get(Activesheet,'Range','B1:B1');
set(ActivesheetRange, 'Value', 'Sx(W/m2)');
ActivesheetRange = get(Activesheet,'Range',['B2:B' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', S_in_xWm(:,1));

ActivesheetRange = get(Activesheet,'Range','C1:C1');
set(ActivesheetRange, 'Value', 'Sy(W/m2)');
ActivesheetRange = get(Activesheet,'Range',['C2:C' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', S_in_yWm(:,1));

ActivesheetRange = get(Activesheet,'Range','D1:D1');
set(ActivesheetRange, 'Value', 'Sz(W/m2)');
ActivesheetRange = get(Activesheet,'Range',['D2:D' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', S_in_zWm(:,1));

ActivesheetRange = get(Activesheet,'Range','E1:E1');
set(ActivesheetRange, 'Value', 'Ex(V/m)');
ActivesheetRange = get(Activesheet,'Range',['E2:E' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', E_in_x(:,1));

ActivesheetRange = get(Activesheet,'Range','F1:F1');
set(ActivesheetRange, 'Value', 'Ey(V/m)');
ActivesheetRange = get(Activesheet,'Range',['F2:F' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', E_in_y(:,1));

ActivesheetRange = get(Activesheet,'Range','G1:G1');
set(ActivesheetRange, 'Value', 'Ez(V/m)');
ActivesheetRange = get(Activesheet,'Range',['G2:G' num2str(number_of_points+1)]);
set(ActivesheetRange, 'Value', E_in_z(:,1));


ActivesheetRange = get(Activesheet,'Range','H1:H1');
set(ActivesheetRange, 'Value', 'Electric Field(V/m)');
ActivesheetRange = get(Activesheet,'Range',['H2:H' num2str(number_of_points+1) ]);
set(ActivesheetRange, 'Value', E_oliko_fi(:,1));

ActivesheetRange = get(Activesheet,'Range','I1:I1');
set(ActivesheetRange, 'Value', 'Sintelestis ekthesis/fi (V/m)');
ActivesheetRange = get(Activesheet,'Range',['I2:I' num2str(number_of_points+1) ]);
set(ActivesheetRange, 'Value', Logos_Hlektrikou_fi);

ActivesheetRange = get(Activesheet,'Range','J1:J1');
set(ActivesheetRange, 'Value', 'Total ElectricField(V/m)');
ActivesheetRange = get(Activesheet,'Range','J2:J2');
set(ActivesheetRange, 'Value', E_total);

ActivesheetRange = get(Activesheet,'Range','K1:K1');
set(ActivesheetRange, 'Value', 'SYNTELESTIS EKTHESIS');
%ActivesheetRange = get(Activesheet,'Range',['K2:K2' num2str(length(Syntelestis_ekthesis))]);
ActivesheetRange = get(Activesheet,'Range','K2:K2');
set(ActivesheetRange, 'Value', Syntelestis_ekthesis);

ActivesheetRange = get(Activesheet,'Range','L1:L1');
set(ActivesheetRange, 'Value', 'S_total_');
ActivesheetRange = get(Activesheet,'Range','L2:L2');
set(ActivesheetRange, 'Value', S_total);


%Now save the workbook
%invoke(Workbook, 'SaveAs', 'myfile.xls'); %to paei sto my documents
invoke(Workbook, 'SaveAs', fullfile(pathname, filename)); %to paei ekei poy epilexame

% To avoid saving the workbook and being prompted to do so,
% uncomment the following code.
 %Workbook.Saved = 1;
 
%% invoke(Workbook, 'Close');
%%
% Quit Excel
invoke(Excel, 'Quit');
%%
% End process
delete(Excel);    

msgbox(['The data were saved at ', fullfile(pathname, filename)]);
plot(handles.axes1, f_i,E_oliko_fi);
xlabel(handles.axes1, 'Frequency [MHz]');
ylabel(handles.axes1, 'Electric Field [Volt/m]');



% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


%setarisma global metavliths se periptwsh selectedindex gia periptwsh
%1,2 ta analoga strings
selectedIndex = get(handles.popupmenu1, 'value');
global limitation
if selectedIndex  == 1
        limitation ='gsm';
elseif selectedIndex == 2;
        limitation = 'dcs';
end

% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.


if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


%setarisma global metavliths se periptwsh selectedindex gia periptwsh
%1,2,3 ta analoga strings
selectedIndex = get(handles.popupmenu2, 'value');
global percent
if selectedIndex  == 1;
        percent ='60';
elseif selectedIndex == 2;
        percent = '70';
elseif selectedIndex == 3;
    percent = '100';
end
    
% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu2


% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in close.
function close_Callback(hObject, eventdata, handles)
% hObject    handle to close (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clc;
clear all;
close Calculate_E;

% --- Executes on button press in estimation.
function estimation_Callback(hObject, eventdata, handles)
% hObject    handle to estimation (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%matlab interpolation with cubic splines at original values of ISOTROPY of PCD8250  factor
