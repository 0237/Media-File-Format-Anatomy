%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% ����
% 1410658
% ��������Ϣ���ļ�����ȡ����
% ע������Ϊһ����������Ϣ��.docx�ļ�
%     ���Ϊһ����ȡ������Ϣ��.txt�ļ�
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% ��.docx���
copyfile('�����ļ�_�������.docx','���ɰ�.zip');
unzip('���ɰ�.zip', '�����ļ����');
delete('���ɰ�.zip');%ɾ�����ɰ�

% ��ȡ������Ϣ
xDoc = xmlread('�����ļ����\word\document.xml');%��ȡ�����.xml�ļ�
wps = xDoc.getElementsByTagName('w:p');%��λ��ǩ
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

% �ָ��ַ�
hexmsg = reshape(hexmsg,2,[])';
hexmsg = uint8(hex2dec(hexmsg))';
hexmsg = native2unicode(hexmsg);

% �����д���ļ�
fid=fopen('��ȡ��Ϣ.txt','w');
fprintf(fid,hexmsg);
fclose(fid);