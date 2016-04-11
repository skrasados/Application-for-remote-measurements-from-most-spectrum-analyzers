
[name,PathName] = uigetfile('*.xls','Select the X-Spectrum Analyzer Trace file');
temp = xlsread(name,1);
EXCELS_File(:,1) = temp(:,1) ;
EXCELS_File(:,2) = temp(:,2);
[name,PathName] = uigetfile('*.xls','Select the Y-Spectrum Analyzer Trace file');
temp = xlsread(name,1);
EXCELS_File(:,3) = temp(:,2);
[name,PathName] = uigetfile('*.xls','Select the Z-Spectrum Analyzer Trace file');
temp = xlsread(name,1);
EXCELS_File(:,4) = temp(:,2);

start_freq=temp(28,3) * 10^(-6)
stop_freq=temp(31,3) * 10^(-6) 
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
P_c = 10 * log10( (bandwidthchannel/bandwidthnoise)* (1/number_of_points) ); 

P_i_x = EXCELS_File(:,2)
P_i_y = EXCELS_File(:,3)
P_i_z = EXCELS_File(:,4)
P_i_x_int = int8(P_i_x)
P_i_y_int = int8(P_i_y)
P_i_z_int = int8(P_i_z)

%apo8hkeysh ston f_i twn syxnothtwn gia tis opoies kaname th metrhsh
% to f_ einai se MHz
f_i = EXCELS_File(:,1)

%f_i = f_i/1000000;


% Ypologismos Isxyos sthn eisodo tou syshmatos(se ka8e stoixeio tou p_out prosti8etai to p_c)
P_out_x = P_i_x + P_c;
P_out_y = P_i_y + P_c;
P_out_z = P_i_z + P_c;  

[File_name,Path_name] = uigetfile('*.xlsx','Select the AntennaFactor and Cable losses file for calculating losses in midrange frequencies')

Isotropic_PCD_corrections_all_table = xlsread('PCD_and_both_cables_data.xlsx');
Isotropic_PCD_corrections_90=Isotropic_PCD_corrections_all_table(801:891,:);
w = start_freq:(stop_freq-start_freq)/(number_of_points-1):stop_freq;

Isotropic_PCD_corrections_A = interp1(Isotropic_PCD_corrections_90(:,2),Isotropic_PCD_corrections_90(:,3),w,'linear');
Isotropic_PCD_corrections_B = interp1(Isotropic_PCD_corrections_90(:,2),Isotropic_PCD_corrections_90(:,4),w,'linear');


Isotropy_dB=(Isotropic_PCD_corrections_A).* f_i' + Isotropic_PCD_corrections_B;
Isotropy_linear = 10.^ (Isotropy_dB/20);

if (start_freq==880) || (start_freq==925) %%gia gsm900
    temp_ = xlsread(File_name);
    Gain_Losses = temp_(801:891,:) ; %sthlh gia sixnothtes 880-970 mhz
else  %%gia gsm1800
    temp_ = xlsread(File_name);
    Gain_Losses = temp_(1620:1820,:) ;  %sthlh gia sixnothtes 1700-1900 mhz
end


Cable_Losses_interp = start_freq:Interpolation_step:stop_freq
CableLosses_interpolated(:,1) = interp1(Gain_Losses(:,2),Gain_Losses(:,9),Cable_Losses_interp,'linear');
display(CableLosses_interpolated);

L_tot = CableLosses_interpolated - Gain_db; 


AF_MHZ_interp = start_freq:Interpolation_step:stop_freq
AF_interpolated(:,1) = interp1(Gain_Losses(:,2),Gain_Losses(:,3),AF_MHZ_interp,'linear')

AF_Absolute = (10 .^ (AF_interpolated(:,1) / 20)) .* Isotropy_linear';


AF_db = 20 * log10(AF_Absolute);
%Antenna_Gain_dB(i) = -29.776613 + 20 * (Log(Frequency_MHz(i)) / Log(10)) - AF_dB(i)
Gain_db = -29.776613 + 20 * log10(f_i) - AF_db;

%
Vsa_dBuV_x = P_i_x + P_c + 106.9897;
Vsa_dBuV_y = P_i_y + P_c + 106.9897;
Vsa_dBuV_z = P_i_z + P_c + 106.9897;

Ebpi_dBuV_ana_m_x = Vsa_dBuV_x + AF_db + CableLosses_interpolated; %AF_dB(i) + Cable_Loss_dB(i)
Ebpi_dBuV_ana_m_y = Vsa_dBuV_y + AF_db + CableLosses_interpolated; %AF_dB(i) + Cable_Loss_dB(i)
Ebpi_dBuV_ana_m_z = Vsa_dBuV_z + AF_db + CableLosses_interpolated; %AF_dB(i) + Cable_Loss_dB(i)

Ebpi_V_ana_m_x = 10 ^ (-6) * 10 .^ (Ebpi_dBuV_ana_m_x / 20);%10 ^ (-6) * 10 ^ (Ebpi_dBuV_ana_m(i) / 20)
Ebpi_V_ana_m_y = 10 ^ (-6) * 10 .^ (Ebpi_dBuV_ana_m_y / 20);%10 ^ (-6) * 10 ^ (Ebpi_dBuV_ana_m(i) / 20)
Ebpi_V_ana_m_z = 10 ^ (-6) * 10 .^ (Ebpi_dBuV_ana_m_z / 20);%10 ^ (-6) * 10 ^ (Ebpi_dBuV_ana_m(i) / 20)

%Ebpi_V_ana_m_tetragono(i) = Ebpi_V_ana_m(i) ^ 2
Ebpi_V_ana_m_tetragono_x = Ebpi_V_ana_m_x .^ 2;
Ebpi_V_ana_m_tetragono_y = Ebpi_V_ana_m_y .^ 2;
Ebpi_V_ana_m_tetragono_z = Ebpi_V_ana_m_z .^ 2;

%Sbpi_W_ana_m2(i) = Ebpi_V_ana_m_tetragono(i) / (120 * 3.1416)
Sbpi_W_ana_m2_x = Ebpi_V_ana_m_tetragono_x / (120 * pi);
Sbpi_W_ana_m2_y = Ebpi_V_ana_m_tetragono_y / (120 * pi);
Sbpi_W_ana_m2_z = Ebpi_V_ana_m_tetragono_z / (120 * pi);

%Ebp = Ebp + Ebpi_V_ana_m_tetragono(i)
%Ebp = Ebp ^ 0.5
Ebp_x=sqrt(sum(Ebpi_V_ana_m_tetragono_x));
Ebp_y=sqrt(sum(Ebpi_V_ana_m_tetragono_y));
Ebp_z=sqrt(sum(Ebpi_V_ana_m_tetragono_z));
%Sbp = Ebp ^ 2 / (120 * 3.1416)
Sbp_x = Ebp_x ^ 2 / (120 * pi);
Sbp_y = Ebp_y ^ 2 / (120 * pi);
Sbp_z = Ebp_z ^ 2 / (120 * pi);

E_total_stratakis=sqrt(Ebp_x ^ 2 + Ebp_y ^ 2 + Ebp_z ^ 2);

S_total_stratakis=sum(Sbpi_W_ana_m2_x+Sbpi_W_ana_m2_y+Sbpi_W_ana_m2_z); %Sbp_x  + Sbp_y  + Sbp_z 

E_oliko_fi = sqrt(Ebpi_V_ana_m_tetragono_x + Ebpi_V_ana_m_tetragono_y + Ebpi_V_ana_m_tetragono_z);

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


Logos_Hlektrikou_fi = E_oliko_fi.^2 ./ (Limit).^2; %logos E_i/E_Limit
Syntelestis_ekthesis = sum(Logos_Hlektrikou_fi);


if (E_oliko_fi < Limit) 
    h = msgbox('Within Limitations','title'); %popup messagebox gia enhmerwsh xrhsth oso aforia an eimaste entos h ektos oriwn
else
    h=msgbox('Out Of Bounds','title','Warn');  %popup messagebox gia enhmerwsh xrhsth oso aforia an eimaste entos h ektos oriwn
end

filename = 'Results.xls';
pathname=pwd;

