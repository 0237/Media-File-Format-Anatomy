%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% ����
% 1410658
% ��������ϢǶ�������ļ�
% ע������Ϊһ��д��������Ϣ��.txt�ļ�
%         ��һ��.docx��ʽ�������ļ�
%     ���Ϊһ����������Ϣ��.docx�ļ�
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% ��.docx���
copyfile('�����ļ�.docx','���ɰ�.zip');
unzip('���ɰ�.zip', '�����ļ����');

% ��ȡ��Ҫ���ص���Ϣ�����ʵ�����
orimsg = char(textread('������Ϣ.txt','%s'));%���ַ���ʽ���ļ�
unimsg = unicode2native(orimsg);%�ַ����ı�����unicodeת��Ϊ����ϵͳ����
hexmsg = dec2hex(unimsg);
msgl = length(hexmsg);
if mod(msgl*2,6) ~= 0 %��ĩβ���0ʹ֮Ϊ6�ı���
    fillchar = repmat('00',3-rem(msgl*2,6)/2+3,1);
    newhexmsg = [hexmsg;fillchar];
end;
newhexmsg = reshape(newhexmsg',1,[]);%ת��Ϊһά����

% Ƕ��������Ϣ
xDoc = xmlread('�����ļ����\word\document.xml');%��ȡ�����.xml�ļ�
wps = xDoc.getElementsByTagName('w:p');%��λ��ǩ
msgl = length(newhexmsg);
for i = 0 : msgl / 6 - 1
    msgseg = newhexmsg(6*i+1:6*(i+1));
    wp = wps.item(i);
    %wp1rsidP = char(wp1.getAttribute('w:rsidP')) 
    wp.setAttribute('w:rsidP',['00',msgseg]);%�޸�wp1��rsidP����
    %wp1rsidP = char(wp1.getAttribute('w:rsidP'))
end;
xmlwrite('�����ļ����\word\document.xml',xDoc);%������������\word\document.xml

% �����.docx�ļ�
dirOutput=dir('�����ļ����');
fileNames={dirOutput.name}';
fileNames(1)=[];%ɾȥ.
fileNames(1)=[];%ɾȥ..
zip('���ɰ�.zip', fileNames, '�����ļ����');
movefile('���ɰ�.zip','�����ļ�_�������.docx');