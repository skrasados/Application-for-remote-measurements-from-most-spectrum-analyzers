function [Start_Frequency,Stop_Frequency,Sweep_Number_Of_Points] = Set_measurement(visaObj,Attenuation,Reference_Level,Start_Frequency,Stop_Frequency,Resolution_BW,Video_BW,Sweep_Number_Of_Points,Sweep_Time,Detector_Function,Trace_Mode,Scale_Type,Number_of_Averages,Instrument_Model)
%global visaObj;
%global Instrument_Model;

%global Attenuation Reference_Level Start_Frequency Stop_Frequency Resolution_BW Video_BW ...
       %Sweep_Number_Of_Points Sweep_Time Detector_Function Trace_Mode Scale_Type ...
       %Number_of_Averages Center_Frequency Date_Time Instrument_Model Instrument_Serial_Number ...
       %Span_Frequency;
   
%Atten,Ref,Res_BW,V_BW,Start,Stop,SP,ST,DT,Sc_Type
%Open_FSH8
instrumentError=0;
%desired settings for the measurement 
%
%set the settings for the dedicated measurement

%%
%set input attenuation
%start with input attenuation a big value e.g. 30dB,and preamplifier off
%return the analyzer in clear Write Trace Mode
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,['input:attenuation ' num2str(Attenuation) ' dB']); %" & Format(Input_Attenuation) & "dB"
else
    %Gia ton E4407B
  fprintf(visaObj,[':SENSe:POWer:RF:ATTenuation ' num2str(Attenuation) ' dB']);
end
%Attenuation
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
%%
%set reference level
fprintf(visaObj,[':DISPlay:WINDow:TRACe:Y:SCALe:RLEVel ' num2str(Reference_Level)]);
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Reference_Level
%%
%set frequency start to 0Hz
fprintf(visaObj,':sense:frequency:start 0');
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
%%
%set stop frequency at the desired value
fprintf(visaObj,[':sense:frequency:stop ' num2str(Stop_Frequency*10^6)]); %" & Format(stop_frequency) '960000000"
fprintf(visaObj,'*WAI');
Stop_Frequency=str2double(query(visaObj,':sense:frequency:stop?'));
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Stop_Frequency
%% 
%set again start frequency at the desired value
fprintf(visaObj,[':sense:frequency:start ' num2str(Start_Frequency*10^6)]); %" & Format(start_frequency) '925000000"
fprintf(visaObj,'*WAI');
Start_Frequency=str2double(query(visaObj,':sense:frequency:start?'));
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Start_Frequency
%%    
%set the resolution bandwidth    
fprintf(visaObj,[':SENSe:BANDwidth:RESolution ' num2str(Resolution_BW*10^3)]);%*10^3)]) %" & Format(RBW) '100000"
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Resolution_BW
%%
%set the video bandwidth
fprintf(visaObj,[':SENSe:BANDwidth:VIDeo ' num2str(Video_BW*10^6)]);%*10^6)]) %" & Format(VBW) '1000000"
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Video_BW
%%
%set the sweep points
%select the model of analyzer first
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    %Sweep_Number_Of_Points=631;
    %fprintf(visaObj,':sense:sweep:points 631');
    Sweep_Number_Of_Points=631
else
    %Gia ton E4407B
    fprintf(visaObj,[':sense:sweep:points ' num2str(Sweep_Number_Of_Points)]);
    %Sweep_Number_Of_Points=str2double(fscanf(visaObj,':SENSe:SWEep:POINts?'));
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj); 
Sweep_Number_Of_Points
%%
%set the sweep time in second
fprintf(visaObj,[':SENSe:SWEep:TIME ' num2str(Sweep_Time) 's']); %" & Format(Sweep_time) '0.2"
%fprintf(visaObj,[':SENSe:SWEep:TIME ' num2str(Sweep_Time)])
%sp=fscanf(visaObj,':sense:sweep:time?')
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Sweep_Time
%%
%Set the Detector_Function_Type to RMS 
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,[':sense:detector:function ' Detector_Function]);
else
    %Gia ton E4407B
   fprintf(visaObj,[':sense:average:type ' Detector_Function]);
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Detector_Function
%%
%set detector to average 
%set the trace mode to average the above number of traces
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,['DISP:WIND:TRAC:MODE ' Trace_Mode]);
else
    %Gia ton E4407B
    fprintf(visaObj,[':sense:detector:function ' Trace_Mode]);
end
%fprintf(visaObj,':sense:detector:function AVERage') %" & Format(Average_Function) 'AVERage"
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Trace_Mode
%%
%set the scale to LOG or LIN
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,[':DISPlay:WINDow:TRACe:Y:SPACing ' Scale_Type])
else
    %Gia ton E4407B
    fprintf(visaObj,[':DISPlay:WINDow:TRACe:Y:SCALe:SPACing ' Scale_Type])
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
Scale_Type
%%
%set the number of trace averages
if strcmp(Instrument_Model,'FSH8');
    %Gia to FSH8
    fprintf(visaObj,[':SENSe:SWEep:COUNt ' num2str(Number_of_Averages)]);
else
    %Gia ton E4407B
    fprintf(visaObj,[':sense:average:count ' num2str(Number_of_Averages)]);
    fprintf(visaObj,[':sense:average:state ON']);
end
fprintf(visaObj,'*WAI');
fprintf(visaObj,'*OPC?');
fscanf(visaObj);
%fprintf(visaObj,'SWE:COUN 100')
%ave=fscanf(visaObj,'SWE:COUN?')
Number_of_Averages
%%
%set the pause interval for matlab halt
% pause_time =Sweep_Time*(1+0.2)*Number_of_Averages+5;
% pause(Sweep_Time*(1+0.2)*Number_of_Averages+5);
%%

%%
%at the end:
%Pause the measurement
%get the trace and the pictrure
%and put them in an Excel file 

%%
%after all:
%return the analyzer in clear Write Trace Mode (clears the average)
%return the analyzer in continous sweep

   
%%


end

