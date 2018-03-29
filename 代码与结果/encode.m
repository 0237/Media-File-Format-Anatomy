%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 杨旭东
% 1410658
% 将隐藏信息嵌入载体文件
% 注：输入为一个写有隐藏信息的.txt文件
%         和一个.docx格式的载体文件
%     输出为一个有隐藏信息的.docx文件
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% 将.docx拆包
copyfile('载体文件.docx','过渡包.zip');
unzip('过渡包.zip', '载体文件拆包');

% 读取需要隐藏的信息并做适当处理
orimsg = char(textread('隐藏信息.txt','%s'));%以字符形式打开文件
unimsg = unicode2native(orimsg);%字符串的编码由unicode转变为本地系统编码
hexmsg = dec2hex(unimsg);
msgl = length(hexmsg);
if mod(msgl*2,6) ~= 0 %在末尾填充0使之为6的倍数
    fillchar = repmat('00',3-rem(msgl*2,6)/2+3,1);
    newhexmsg = [hexmsg;fillchar];
end;
newhexmsg = reshape(newhexmsg',1,[]);%转化为一维数组

% 嵌入秘密信息
xDoc = xmlread('载体文件拆包\word\document.xml');%读取载体的.xml文件
wps = xDoc.getElementsByTagName('w:p');%定位标签
msgl = length(newhexmsg);
for i = 0 : msgl / 6 - 1
    msgseg = newhexmsg(6*i+1:6*(i+1));
    wp = wps.item(i);
    %wp1rsidP = char(wp1.getAttribute('w:rsidP')) 
    wp.setAttribute('w:rsidP',['00',msgseg]);%修改wp1的rsidP属性
    %wp1rsidP = char(wp1.getAttribute('w:rsidP'))
end;
xmlwrite('载体文件拆包\word\document.xml',xDoc);%保存结果到测试\word\document.xml

% 打包成.docx文件
dirOutput=dir('载体文件拆包');
fileNames={dirOutput.name}';
fileNames(1)=[];%删去.
fileNames(1)=[];%删去..
zip('过渡包.zip', fileNames, '载体文件拆包');
movefile('过渡包.zip','载体文件_载入完成.docx');