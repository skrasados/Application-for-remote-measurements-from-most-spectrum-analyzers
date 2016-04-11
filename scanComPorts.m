function ports = scanComPorts
% SCANCOMPORTS - Scan for available serial ports. 
% The function returns a cell structure with the available serial ports.
% 
% Example:
% p = scanComPorts;
%
% Known Issues: Matlab sometimes cannot see an open port and requires a restart (not shortcoming of the code).
%
% Developed by Chandrakanth R Terupally | www.matlabtraining.com
%

ss = serial('COM129');
try fopen(ss)
catch err
end
p = regexp(err.message, 'COM[0-9]*', 'match');
p = p(2:end);
ports = cell(0,1);
for i = 1:size(p,2)
    s = serial(p{i});
    try fopen(s)
        ports(end+1) = p(i);
        fclose(s);
    catch err
        continue
    end
end
