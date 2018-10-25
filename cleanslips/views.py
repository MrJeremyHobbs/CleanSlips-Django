from django.shortcuts import render, redirect
from django.http import HttpResponseBadRequest, HttpResponse
from django import forms
#import django_excel as excel

# modules
import os
import re
from operator import itemgetter
from io import StringIO
from io import BytesIO

# local modules
from . modules import callnumber
from . modules.docx_mailmerge_local.mailmerge import MailMerge

# Form class
class UploadFileForm(forms.Form):
    file = forms.FileField()


# Create your views here.
def upload(request, campus):     
    
    # check campus code
    campuses = []
    
    campus = campus.upper()
    
    campus_name = ""
    if campus == "CBA":
        campus_name = "Bakersfield"

    if campus == "U$C":
        campus_name = "Channel Islands"

    if campus == "CCH":
        campus_name = "Bakersfield"

    if campus == "CDH":
        campus_name = "Bakersfield"

    if campus == "CSH":
        campus_name = "Bakersfield"

    if campus == "CFS":
        campus_name = "Bakersfield"

    if campus == "CFI":
        campus_name = "Bakersfield"

    if campus == "CHU":
        campus_name = "Bakersfield"

    if campus == "CLO":
        campus_name = "Bakersfield"

    if campus == "CLA":
        campus_name = "Bakersfield"

    if campus == "CVM":
        campus_name = "Bakersfield"

    if campus == "MB@":
        campus_name = "Bakersfield"

    if campus == "MFL":
        campus_name = "Bakersfield"

    if campus == "CNO":
        campus_name = "Bakersfield"

    if campus == "CPO":
        campus_name = "Bakersfield"

    if campus == "CSA":
        campus_name = "Bakersfield"

    if campus == "CSB":
        campus_name = "Bakersfield"
    
    if campus == "CDS":
        campus_name = "Bakersfield"

    if campus == "CSF":
        campus_name = "Bakersfield"

    if campus == "CSJ":
        campus_name = "Bakersfield"

    if campus == "CPS":
        campus_name = "Bakersfield"

    if campus == "CS1":
        campus_name = "Bakersfield"

    if campus == "CSO":
        campus_name = "Bakersfield"

    if campus == "CTU":
        campus_name = "Bakersfield"

    if request.method == "POST":
        form = UploadFileForm(request.POST, request.FILES)
        
        if form.is_valid():
            filehandle = request.FILES['file']
            
            # read Alma spreadsheet
            ill_requests = []
            
            rows = filehandle.get_array()
            for row in rows:
                title = row[0]
                
                # skip header
                if title == "Title":
                    continue
                
                author = row[1]
                publisher = row[2]
                publication_date = row[3]
                barcode = row[4]
                isbn_issn = row[5]
                availability = row[6]
                volume_issue = row[7]
                requestor_email = row[9]
                pickup_at = row[10]
                electronic_available = row[11]
                digital_available = row[12]
                external_request_id = row[13]
                partner_name = row[14]
                partner_code = row[15]
                copyright_status = row[16]
                level_of_service = row[17]
                
                # parse shipping note
                shipping_note = row[8]
                shipping_notes = shipping_note.split('||')
                comments = shipping_notes[0]
                requestor_name = shipping_notes[1]
                
                # parse availability
                regex = r'(.*?),(.*?)\.(.*).*(\(\d{1,3} copy,\d{1,3} available\))'
                q = re.findall(regex, availability)
                matches = list(q[0])
                
                library = matches[0]
                location = matches[1]
                call_number = matches[2]
                holdings = matches[3]
                
                full_availability = f"[{location} - {call_number[:-1]}]" # negative index to remove extra space
                
                # normalize call number for sorting
                lccn = callnumber.LC(call_number)
                lccn_components = lccn.components(include_blanks=True)
                normalized_call_number = lccn.normalized
                if normalized_call_number == None:
                    normalized_call_number = call_number
                     
                sort_string = f"{normalized_call_number}|{location}"

                # add to requests dictionary
                ill_request = {
                    'Partner_name' : partner_name,
                    'External_request_ID' : external_request_id,
                    'Availability' : full_availability,
                    'Call_Number' : call_number,
                    'Comments' : comments,
                    'RequestorName' : requestor_name,
                    'VolumeIssue' : volume_issue,
                    'Title' : title,
                    'Shipping_note' : requestor_name,
                    'Sort' : sort_string,
                }
                
                # add to ongoing list
                ill_requests.append(ill_request)   
                
            # sort requests by location and normalized call number
            requests_sorted = sorted(ill_requests, key=itemgetter('Sort'))
                
            # mail merge
            template = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'templates/labels/TEMPLATE_stickers.docx')

            document = MailMerge(template)
            if "_slips" in template:
                document.merge_templates(requests_sorted, separator='nextColumn_section')
            if "_stickers" in template:
                document.merge_rows('Shipping_note', requests_sorted)
            
            f = BytesIO()
            document.write(f)
            length = f.tell()
            f.seek(0)
            response = HttpResponse(
                f.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            response['Content-Disposition'] = 'attachment; filename=SLIPS.docx'
            response['Content-Length'] = length
            return response
    
    else:
        form = UploadFileForm()
    return render(
        request,
        'upload_form.html',
        {
            'form': form,
            'title': 'CleanSlips',
            'header': (f'CleanSlips')
        })
        
        
def get_campus_name ():
    pass