function [f_MHz_Interpolated,Data_Interpolated]=linear_interpolation(start_freq,stop_freq,number_of_points)
%example [x,y=linear_interpolation(2400,2500,631)
%returns x=f_MHz_Interpolated se pinaka stili, y=Data_Interpolated se pinaka stili gia 631 times 
%anamesa stis syxnotites 2400 eos 2500

clc;

%start_freq=2400; %gia parametropoisi gia alli syxnotita ayti i entoli feygei
%stop_freq=2500; %gia parametropoisi gia alli syxnotita ayti i entoli feygei
%number_of_points=631; %gia parametropoisi gia allo arithmo sweep points ayti i entoli feygei

[File_name,Path_name] = uigetfile('*.xlsx','Select the .xlsx file for linear interference'); %arxeio me Antena Factor Data i arxeio me cables attenuation
full_file_name=fullfile(Path_name,File_name);
All_per_MHz_Data = xlsread(full_file_name,1); % i xlsread epistrefei mono ta numeric data num = xlsread(filename,sheet) reads the specified worksheet.

%the fisrt frequency in the table is the above
first_All_per_MHz_Data_Frequency=All_per_MHz_Data(1,1); % matlab variable (mia grammi ligoteri giati den lambanei ypocin tis kafalides stin proti grammi toy Excel)
from_number_record=start_freq-first_All_per_MHz_Data_Frequency+1; % matlab variable (mia grammi ligoteri giati den lambanei ypocin tis kafalides stin proti grammi toy Excel)
to_number_record=stop_freq-first_All_per_MHz_Data_Frequency+1; % matlab variable (mia grammi ligoteri giati den lambanei ypocin tis kafalides stin proti grammi toy Excel)

%Get from this table the data between Start and stop frequency
Data_to_interpolate=All_per_MHz_Data(from_number_record:to_number_record,:); %briskoyme tis sostes grammes (eggrafes) sto arxeio poy na symbadizoyn me tis star kai stop frequencies

%briskoyme ton pinaka tis parembolis syxnotiton se MHz
%to step gia na mas dosei paremblimenes times ises me ton ariumo ton
%simeion poy theloyme anamesa stin arxiki kai stin teliki syxnotita einai:
Interpolation_step=(stop_freq-start_freq)/(number_of_points-1);
f_MHz_Interpolated = start_freq:Interpolation_step:stop_freq; % frequencies table in MHz interpolated (pinakas grammi)
f_MHz_Interpolated=f_MHz_Interpolated' %pinakas stili

%briskoyme ta antistoixa dedomena parembolis grammikis
%yi = interp1(x,Y,xi,method) interpolates using alternative methods
Data_Interpolated = interp1(Data_to_interpolate(:,1),Data_to_interpolate(:,2),f_MHz_Interpolated,'linear'); %pinakas grammi
%Data_Interpolated=Data_Interpolated' %pinakas stili
%AF=Data_Interpolated(1,:);
%Af1=Data_Interpolated;
end