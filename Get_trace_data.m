function [Trace_data]=Get_trace_data(visaObj,Instrument_Model)
%global Attenuation Reference_Level Start_Frequency Stop_Frequency Resolution_BW Video_BW ...
       %Sweep_Number_Of_Points Sweep_Time Detector_Function Trace_Mode Scale_Type ...
       %Number_of_Averages Center_Frequency Date_Time Instrument_Model Instrument_Serial_Number ...
       %Span_Frequency;
   %UNTITLED3 Summary of this function goes here

if strcmp(Instrument_Model,'FSH8')
    %Gia to FSH8 pernei ta dedomena se bit
    fprintf(visaObj,'FORM:DATA REAL,32');
    fprintf(visaObj,':TRACE:DATA? TRACE1');
    pause(0.2)
    Trace_data = binblockread(visaObj,'float32');
       
else
    %Gia ton E4407B pernei ta dedomena se ascii code
    fprintf(visaObj,':FORMat:TRACe:DATA ASCii');
    fprintf(visaObj,':TRACE:DATA? TRACE1');
    Trace_data = fscanf(visaObj,':TRACE:DATA? TRACE1');
    fid=fopen('Last_trace.txt','wt');
    fprintf(fid,Trace_data);

    fclose(fid);
    Trace_data=importdata('Last_trace.txt');
    Trace_data=Trace_data';
    %delete('Last_trace.txt');
    
end
%fprintf(visaObj,'FORM ASC');gia ASCII format of received data
%fprintf(visaObj,'FORM:DATA REAL,32');%;gia REAL,32 format of received data
%In REAL,32 format, a string of return values would look like: 
%#42524<value 1><value 2>...<value n> 
%with
%#4  Number of digits of the following number of data bytes (= 4 in this example)
%2524  Number of following data bytes (2524, corresponds to the 631 sweep points of the 
%R&S FSH.
%<value>  4-byte floating point value
%%
%fscanf(visaObj,'FORM:DATA?') den paixei ayth
%datatype=fscanf(visaObj)

%fprintf(visaObj,':INITiate:CONTinuous OFF;*WAI')
%pause(5)
%%
%fprintf(visaObj,':TRACE:DATA? TRACE1');
%%
%%works and gives count =2530=2524+6 where 6 characters are #42524 for 631
%%points of FSH8
%A is a collumn vector of the trace of FSH8
%[A,count] = binblockread(visaObj, 'uint8')
%%
%use this command which gives the Trace values in one column
   fprintf(visaObj,'FORM:DATA REAL,32');
    fprintf(visaObj,':TRACE:DATA? TRACE1');
    pause(0.2)
    Trace_data = binblockread(visaObj,'float32');
% 


end

