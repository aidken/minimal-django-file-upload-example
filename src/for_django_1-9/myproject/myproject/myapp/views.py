# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.template import RequestContext
from django.http import HttpResponse, HttpResponseRedirect
from django.core.urlresolvers import reverse

from myproject.myapp.models import Document
from myproject.myapp.forms import DocumentForm

import xlsxwriter


def list(request):
    # Handle file upload
    if request.method == 'POST':
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():

            # # original code
            # newdoc = Document(docfile=request.FILES['docfile'])
            # newdoc.save()

            # # Redirect to the document list after POST
            # return HttpResponseRedirect(reverse('list'))
            # # orignal code up to here commented out

            # set up an Excel file
            # cf: https://xlsxwriter.readthedocs.org/en/latest/example_http_server3.html
            import io
            output    = io.BytesIO()
            workbook  = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet()

            # do something to the file being uploaded
            for row, item in enumerate(request.FILES['docfile']):
                # logging.debug(item)
                # logging.debug( type(item) )
                worksheet.write( row, 0, item.decode() )  # item is bytes. convert it to string

            workbook.close()
            excel_data = output.getvalue()

            # Rewind the buffer.
            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Report.xlsx'
            response.write(excel_data)

            # # actually I don't need save the original file
            # newdoc = Document(docfile=request.FILES['docfile'])
            # newdoc.save()
            # logging.debug('type of newdoc is {}.'.format(type(newdoc)) )
            # logging.debug(newdoc)

            # # Redirect to the document list after POST
            # return HttpResponseRedirect(reverse('myproject.myapp.views.list'))
            return response

    else:
        form = DocumentForm()  # A empty, unbound form

    # Load documents for the list page
    documents = Document.objects.all()

    # Render list page with the documents and the form
    return render(
        request,
        'list.html',
        {'documents': documents, 'form': form}
    )
