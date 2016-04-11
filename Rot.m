function [x] = Rot(direction,port)
h = msgbox('Wait For Antenna Position');
%s = serial('COM15');
s= serial(port)
%set(s,'BaudRate',9600);
try
fopen(s);
reply = fscanf(s);
reply = fscanf(s);
if strfind(reply, 'System Ready') ==1 
    x=1;
else
    x=0;
end
 
fprintf(s,direction);
direction
reply = fscanf(s);
   
fclose(s);
delete(s);
clear srepl;

catch
  try
       fclose(s);
  catch
      delete(s);
      
  end
  clear srepl;
end
delete(h);