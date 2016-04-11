function spic2xls(Excel,visaObj,Instrument_Model,pic,file,sheet,param)

[fpath,fname,fext] = fileparts(file);
if isempty(fpath);
    out_path = pwd;
elseif fpath(1)=='.';
    out_path = [pwd filesep fpath];
else
    out_path = fpath;
end
%

% Activating Sheet
Sheets = Excel.Worksheets;
sheet = get(Sheets, 'Item', sheet);
invoke(sheet, 'Activate');
Activesheet = Excel.Activesheet;

%% Adding Picture
if strcmp(Instrument_Model,'FSH8');
    
    %Gia to FSH8
    
    pic='FSH8.png';
    fprintf(visaObj,':DISPLAY:WIND:FETC?');% PNG, SCREEN, GRAYSCALE');
    FSH8_PNG = binblockread(visaObj,'uint8'); fread(visaObj,1);
    % save as a PNG  file in the current matlab directory
    fid = fopen('FSH8.png','w');
    fwrite(fid,FSH8_PNG,'uint8');
    fclose(fid);
    % Read the PNG and display image
    figure; colormap(gray(256)); 
    imageMatrix = imread('FSH8.png','png');
    image(imageMatrix);
    %delete FSH8.png
    % Adjust the figure so it shows accurately
    sizeImg = size(imageMatrix);
    set(gca,'Position',[0 0 1 1],'XTick' ,[],'YTick',[]); set(gcf,'Position',[50 50 640 480]);
    axis off; axis image;
    %Excel.Activesheet.UsedRange = get(Activesheet,'Range',['A1:A1'])
    %print -dmeta
   
else
    
    %if Instrument_Model=='E4407'
    %Gia ton E4407B
    
    pic='E4407scr.gif';
    fprintf(visaObj,'*CLS');
    % Store the current screen image in the file "E4407gif" to the analyzer R disk
    %
    fprintf(visaObj, 'MMEM:STOR:SCR "R:STscreen.gif"');

    % Transfer the image to MATLAB
    fprintf(visaObj, 'MMEM:DATA? "R:STscreen.gif"');

    E4407gif = binblockread(visaObj,'uint8'); 
    fread(visaObj,1);
    %%
    %delete the screen image from analyzer
    fprintf(visaObj, 'MMEM:DEL "R:STscreen.gif"')
    %fprintf(visaObj,'*CLS')
    % save as a PNG  file in the current matlab directory
    fid = fopen('E4407scr.gif','w');
    fwrite(fid,E4407gif,'uint8');
    fclose(fid);
    figure; %colormap(gray(15));
    imageMatrix = imread('E4407scr.gif','gif');
    image(imageMatrix); 
    % Adjust the figure so it shows accurately
    sizeImg = size(imageMatrix);
    %set(gca,'Position',[0 0 1 1],'XTick' ,[],'YTick',[]); set(gcf,'Position',[50 50 sizeImg(2) sizeImg(1)]);
    %gia na ginei 640x480 pixels dinoyme
    set(gca,'Position',[0 0 1 1],'XTick' ,[],'YTick',[]); set(gcf,'Position',[50 50 640 480]);
    colormap vga
    axis off; axis image;
end

%instrumentError = query(visaObj,':SYSTEM:ERR?');
%while ~isequal(instrumentError,['+0,"No error"' char(10)])
    %disp(['Instrument Error: ' instrumentError]);
    %instrumentError = query(visaObj,':SYSTEM:ERR?');
%end

[ifpath,ifname,ifext] = fileparts(pic);
if isempty(ifpath);
    iout_path = pwd;
elseif fpath(1)=='.';
    iout_path = [pwd filesep ifpath];
else
    iout_path = ifpath;
end
% send command and get PNG.

% Function AddPicture(Filename As String, LinkToFile As MsoTriState,
% SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single) As Shape
%
ExAct = Excel.Activesheet;


invoke(ExAct.Shapes,'AddPicture',[iout_path filesep ifname ifext],0,1,0,0,640,480);%pixels
png_file=[iout_path filesep ifname ifext];
%points to pixels http://www.endmemo.com/sconvert/pixelpoint.php
%delete [iout_path filesep ifname ifext]
%cm to pixels http://www.unitconversion.org/typography/centimeters-to-pixels-y-conversion.html
