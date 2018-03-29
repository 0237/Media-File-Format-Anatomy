%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 杨旭东
% 1410658
% 将隐藏信息从文件中提取出来
% 注：输入为一个有隐藏信息的.docx文件
%     输出为一个提取隐藏信息的.txt文件
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% 将.docx拆包
copyfile('载体文件_载入完成.docx','过渡包.zip');
unzip('过渡包.zip', '载体文件拆包');
delete('过渡包.zip');%删除过渡包

% 提取秘密信息
xDoc = xmlread('载体文件拆包\word\document.xml');%读取载体的.xml文件
wps = xDoc.getElementsByTagName('w:p');%定位标签
hexmsg = '';i=0;
while true
    wp = wps.item(i);
    wprsidP = char(wp.getAttribute('w:rsidP'));
    if strcmp(wprsidP,'00000000') == 1 || i == wps.getLength()-1
        break;
    end;
    hexmsg = [hexmsg,wprsidP(3:8)];
    i=i+1;
end;

% 恢复字符
hexmsg = reshape(hexmsg,2,[])';
hexmsg = uint8(hex2dec(hexmsg))';
hexmsg = native2unicode(hexmsg);

% 将结果写入文件
fid=fopen('提取信息.txt','w');
fprintf(fid,hexmsg);
fclose(fid);