function [ output_args ] = plot_SA_Trace(visaObj, Trace_data) 

%get Reference level for plotting purposes
ref_lev=str2num(query(visaObj,':DISP:WIND:TRAC:Y:RLEV?'));
Start_Frequency=str2num(query(visaObj,':SENSe:FREQuency:STARt?'))/10^6; %in MHz
Stop_Frequency=str2num(query(visaObj,':SENSe:FREQuency:STOP?'))/10^6; %in MHz
%%
%remove the terminator character
%fscanf(visaObj)
%%
%diafora metaxi binblockread kai fread
%i binblockread epistr;efei ston pinaka A mono ta bytes toy Trace
%i fread epistrefei kai ta arxika #42524 kai ton teliko xaraktira chr$(10)
%fprintf(visaObj,':TRACE:DATA? TRACE1');
%[B,count] = fread(visaObj)%, 'uint8')
%%
%Trace1_Data=fscanf(visaObj,':TRACE:DATA? TRACE1')
%Put analyzer in continous mode after getting the trace
%fprintf(visaObj,':INITiate:CONTinuous ON;*WAI')
%pause(5)
%sPut2Excel(visaObj)


%%
% create and bring to front figure number 1 
%figure(1); 
nr_points=size(Trace_data,1)
% Plot trace data vs sweep point index 
%plot(1:nr_points,data) 

%kanei plot ta dedomena tou pinaka table se sxesh me ta bit pou pernei apo
%ton analith se bit binblockread(visaObj,'float32');
frequency_table=Start_Frequency:(Stop_Frequency-Start_Frequency)/(nr_points-1):Stop_Frequency;

plot(frequency_table,Trace_data)
% Adjust the x limits to the nr of points 
% and the y limits for 100 dB of dynamic range 
%xlim([1 nr_points])
xlim([Start_Frequency Stop_Frequency]);
ylim([ref_lev-100 ref_lev]);
% activate the grid lines 
grid on 
% xlabel('Frequency in MHz'); 
% ylabel('Amplitude (dBm)'); 
end

