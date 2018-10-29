# -*- encoding: utf-8 -*-

from email.header import decode_header

# 提取邮件附件
def get_email_cntent(message, base_save_path):
    j = 0
    content = ''
    attachment_files = []
    for part in message.walk():
        j = j + 1
        file_name = part.get_filename()
        contentType = part.get_content_type()
        # 保存附件
        if file_name:  # Attachment
            # Decode filename
            # print(file_name)
            # h = email.Header.Header(file_name)
            # dh = email.Header.decode_header(h)
            dh = decode_header(file_name)
            if dh[0][1]:  # 如果包含编码的格式，则按照该格式解码
                # filename = unicode(filename, dh[0][1])
                filename = str(dh[0][0], encoding="gb2312")
                print("获取到附件:"+filename)
            data = part.get_payload(decode=True)
            att_file = open(base_save_path + filename, 'wb') # base_save_path 保存当前路径
            attachment_files.append(filename)
            att_file.write(data)
            att_file.close()
            print("保存成功!路径----->"+base_save_path)
        # elif contentType == 'text/plain' or contentType == 'text/html':
        #     # 保存正文
        #     data = part.get_payload(decode=True)
        #     charset = guess_charset(part)
        #     if charset:
        #         charset = charset.strip().split(';')[0]
        #         print ('charset:', charset)
        #         data = data.decode(charset)
        #     content = data
    return content, attachment_files


def get_email_headers(msg):
    # 邮件的From, To, Subject存在于根对象上:
    headers = {}
    for header in ['From', 'To', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['date'] = value
            if header == 'Subject':
                # 需要解码Subject字符串:
                subject = decode_str(value)
                headers['subject'] = subject
            else:
                # 需要解码Email地址:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = u'%s <%s>' % (name, addr)
                if header == 'From':
                    from_address = value
                    headers['from'] = from_address
                else:
                    to_address = value
                    headers['to'] = to_address
    content_type = msg.get_content_type()
    print ('head content_type: ', content_type)
    return headers



def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value