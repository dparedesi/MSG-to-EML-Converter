# msg_converter_core.py

import extract_msg
import os
import re
import email
from email.message import Message as EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.message import MIMEMessage
from email.header import Header
from email import encoders
from email import utils as email_utils
import mimetypes
import time

# --- Helper Functions (sanitize_filename, guess_mimetype) ---
def sanitize_filename(filename_str, default_name="unnamed_file"):
    if not filename_str:
        return default_name
    filename_str = os.path.basename(filename_str)
    s = re.sub(r'[\\/*?:"<>|]', "_", filename_str)
    s = re.sub(r'[\s_]+', "_", s)
    s = s.strip('_')
    if not s:
        return default_name
    return s

def guess_mimetype(filename):
    m_type, encoding = mimetypes.guess_type(filename)
    if m_type:
        return m_type.split('/', 1)
    return 'application', 'octet-stream'

# --- Core EML Building Logic ---
def build_eml_from_msg_recursively(msg_instance: extract_msg.Message, 
                                   nesting_level=0, 
                                   log_callback=print) -> EmailMessage: # Added log_callback
    subject_for_log = getattr(msg_instance, 'subject', 'N/A')
    log_callback(f"{'  ' * nesting_level}Building EML for MSG (Subject: '{subject_for_log}')")
    
    eml_obj = EmailMessage() 

    # === 1. Set Headers ===
    if getattr(msg_instance, 'subject', None):
        eml_obj['Subject'] = Header(msg_instance.subject, 'utf-8').encode()
    
    sender_val = getattr(msg_instance, 'sender', None)
    if sender_val:
        eml_obj['From'] = sender_val

    to_val = getattr(msg_instance, 'to', None)
    if to_val:
        eml_obj['To'] = to_val

    cc_val = getattr(msg_instance, 'cc', None)
    if cc_val:
        eml_obj['Cc'] = cc_val

    msg_parsed_date_val = getattr(msg_instance, 'parsedDate', None)
    date_set_successfully = False
    if msg_parsed_date_val:
        if hasattr(msg_parsed_date_val, 'timetuple') and callable(getattr(msg_parsed_date_val, 'timetuple')):
            try:
                eml_obj['Date'] = email_utils.format_datetime(msg_parsed_date_val)
                date_set_successfully = True
            except Exception as e_fmt_dt:
                 log_callback(f"{'  ' * nesting_level}  WARNING: Error formatting datetime object {msg_parsed_date_val} with format_datetime: {e_fmt_dt}")
        elif isinstance(msg_parsed_date_val, tuple) and len(msg_parsed_date_val) >= 6:
            try:
                time_tuple_for_mktime = list(msg_parsed_date_val[:9])
                while len(time_tuple_for_mktime) < 9: 
                    if len(time_tuple_for_mktime) >= 6: 
                         time_tuple_for_mktime.append(0 if len(time_tuple_for_mktime) < 8 else -1)
                    else: break 
                
                if len(time_tuple_for_mktime) == 9:
                    timestamp = time.mktime(tuple(time_tuple_for_mktime))
                    eml_obj['Date'] = email_utils.formatdate(timestamp, localtime=True)
                    date_set_successfully = True
                else:
                    log_callback(f"{'  ' * nesting_level}  WARNING: parsedDate tuple {msg_parsed_date_val} could not be reliably formed into a 9-element tuple for mktime.")
            except (TypeError, ValueError, OverflowError) as e_tuple_conv:
                log_callback(f"{'  ' * nesting_level}  WARNING: Could not convert parsedDate tuple {msg_parsed_date_val} to valid Date header: {e_tuple_conv}")
    
    if not date_set_successfully and getattr(msg_instance, 'date', None):
        eml_obj['Date'] = msg_instance.date

    message_id_val = getattr(msg_instance, 'messageId', None)
    if message_id_val:
        eml_obj['Message-ID'] = message_id_val
    
    if not eml_obj['MIME-Version']:
        eml_obj['MIME-Version'] = '1.0'

    # === 2. Prepare Body Part(s) ===
    body_structure_parts = [] 
    plain_body = getattr(msg_instance, 'body', None)
    html_body = getattr(msg_instance, 'htmlBody', None)

    if html_body and plain_body:
        alt_multipart = MIMEMultipart('alternative')
        alt_multipart.attach(MIMEText(plain_body, 'plain', _charset='utf-8'))
        alt_multipart.attach(MIMEText(html_body, 'html', _charset='utf-8'))
        body_structure_parts.append(alt_multipart)
    elif html_body:
        body_structure_parts.append(MIMEText(html_body, 'html', _charset='utf-8'))
    elif plain_body:
        body_structure_parts.append(MIMEText(plain_body, 'plain', _charset='utf-8'))
    else:
        body_structure_parts.append(MIMEText('', 'plain', _charset='utf-8'))

    # === 3. Prepare "File" Attachment Parts ===
    file_attachment_mime_parts = []
    for att_index, att in enumerate(getattr(msg_instance, 'attachments', [])):
        att_long_filename = getattr(att, 'longFilename', None)
        att_short_filename = getattr(att, 'shortFilename', None)
        default_att_name = f"attachment_{att_index + 1}"
        att_name_for_disposition = sanitize_filename(att_long_filename or att_short_filename or default_att_name, default_att_name)

        if isinstance(att.data, extract_msg.Message):
            log_callback(f"{'  ' * (nesting_level + 1)}-> Nested MSG: '{att_name_for_disposition}'. Recursively building EML...")
            nested_eml_obj = build_eml_from_msg_recursively(att.data, nesting_level + 1, log_callback) # Pass callback
            if nested_eml_obj:
                mime_message_part = MIMEMessage(nested_eml_obj)
                nested_subject = getattr(att.data, 'subject', None)
                eml_att_filename = sanitize_filename(f"{nested_subject or 'NestedMessage'}.eml", "NestedMessage.eml")
                mime_message_part.add_header('Content-Disposition', f'attachment; filename="{eml_att_filename}"')
                file_attachment_mime_parts.append(mime_message_part)
                log_callback(f"{'  ' * (nesting_level + 1)}  Done building nested EML part: {eml_att_filename}")
        else: 
            if hasattr(att, 'data') and att.data:
                log_callback(f"{'  ' * (nesting_level + 1)}-> Regular Attachment: '{att_name_for_disposition}'")
                maintype, subtype = guess_mimetype(att_long_filename or att_short_filename or "")
                
                if maintype == 'text':
                    try:
                        att_payload_str = att.data.decode('utf-8')
                        reg_att_part = MIMEText(att_payload_str, subtype, _charset='utf-8')
                    except UnicodeDecodeError:
                        reg_att_part = MIMEBase(maintype, subtype, name=att_name_for_disposition)
                        reg_att_part.set_payload(att.data)
                        encoders.encode_base64(reg_att_part)
                else:
                    reg_att_part = MIMEBase(maintype, subtype, name=att_name_for_disposition)
                    reg_att_part.set_payload(att.data)
                    encoders.encode_base64(reg_att_part)
                
                reg_att_part.add_header('Content-Disposition', f'attachment; filename="{att_name_for_disposition}"')
                file_attachment_mime_parts.append(reg_att_part)

    # === 4. Assemble final EML object structure ===
    if not eml_obj['MIME-Version']: # Ensure it's set before payload/attach
        eml_obj['MIME-Version'] = '1.0'
        
    the_body_structure_part = body_structure_parts[0] 

    if not file_attachment_mime_parts:
        eml_obj.set_payload(the_body_structure_part.get_payload())                                                                
        for k_hdr in list(eml_obj.keys()): 
            if k_hdr.lower().startswith('content-'):
                del eml_obj[k_hdr] 
        for k_hdr, v_hdr in the_body_structure_part.items():
            if k_hdr.lower().startswith('content-'): 
                eml_obj[k_hdr] = v_hdr 
        if not eml_obj['Content-Type'] and hasattr(the_body_structure_part, 'get_content_type'):
            eml_obj.set_type(the_body_structure_part.get_content_type())
    else:
        if 'Content-Type' in eml_obj and not eml_obj.is_multipart():
            del eml_obj['Content-Type']
        eml_obj.attach(the_body_structure_part)
        for part in file_attachment_mime_parts:
            eml_obj.attach(part)
        current_content_type = eml_obj.get_content_type()
        if not current_content_type.startswith('multipart/mixed'):
            payload_parts = eml_obj.get_payload() 
            eml_obj.set_payload([]) 
            eml_obj.set_type('multipart/mixed')
            for p_item in payload_parts:
                if isinstance(p_item, EmailMessage): 
                    eml_obj.attach(p_item)
                else:
                    log_callback(f"{'  ' * nesting_level}  WARNING: Non-Message item in payload list during multipart/mixed reconstruction: {type(p_item)}")
    return eml_obj

# --- Main Conversion Function to be called by Streamlit app ---
def convert_msg_to_single_eml(msg_file_path_or_bytes, output_eml_filename_stem, log_callback=print):
    """
    Converts a .msg file (from path or bytes) to a single .eml (bytes).
    output_eml_filename_stem is used for the suggested download name.
    Returns (eml_bytes, suggested_download_filename) or (None, None) on failure.
    """
    log_callback(f"Starting conversion of main MSG: '{output_eml_filename_stem}'")
    try:
        # extract_msg.Message can take a path or a file-like object (BytesIO for bytes)
        main_msg_instance = extract_msg.Message(msg_file_path_or_bytes)
    except Exception as e:
        log_callback(f"Error reading main MSG file: {e}")
        return None, None

    final_eml_message_obj = build_eml_from_msg_recursively(main_msg_instance, log_callback=log_callback)

    if final_eml_message_obj:
        try:
            eml_bytes = final_eml_message_obj.as_bytes()
            suggested_filename = f"{sanitize_filename(output_eml_filename_stem)}.eml"
            log_callback(f"Successfully converted MSG to EML bytes. Suggested filename: {suggested_filename}")
            return eml_bytes, suggested_filename
        except Exception as e:
            log_callback(f"Error serializing final EML object to bytes: {e}")
            return None, None
    else:
        log_callback("Failed to produce a final EML object from the main MSG file.")
        return None, None

if __name__ == '__main__':
    # This part is for testing the core logic directly, not used by Streamlit app
    print("Testing msg_converter_core.py directly...")
    test_msg_file = "Name_of_email.msg" # Replace with your test .msg file
    if os.path.exists(test_msg_file):
        eml_content_bytes, download_name = convert_msg_to_single_eml(test_msg_file, os.path.splitext(os.path.basename(test_msg_file))[0])
        if eml_content_bytes:
            output_test_dir = "test_eml_output_core"
            if not os.path.exists(output_test_dir):
                os.makedirs(output_test_dir)
            output_path = os.path.join(output_test_dir, download_name)
            with open(output_path, 'wb') as f:
                f.write(eml_content_bytes)
            print(f"Test EML saved to: {output_path}")
        else:
            print("Test conversion failed.")
    else:
        print(f"Test MSG file not found: {test_msg_file}")