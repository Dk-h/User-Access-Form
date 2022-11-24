import shutil
import PySimpleGUI as gui
import os
from docxtpl import DocxTemplate
import docx2pdf
import subprocess

# context to be rendered in word file
context = {}

# gui.theme("DarkGrey3")
# gui.theme("Topanga")
gui.theme("DarkBrown5")
# gui.theme_previewer()


def internetAccess1():

    header = [
        gui.Text('IT Access required', size=(20, 2), justification='c'),
        gui.Text('Date till which access is required',
                 size=(20, 2), justification='c'),
        gui.Text('Additional approvals required to be enclosed',
                 size=(20, 2), justification='c'),
        gui.Text('Additional Information /Remarks',
                 size=(20, 2), justification='c')
    ]
    row = []
    for i in range(0, 5):
        col = []
        for j in range(0, 3):
            if i == 1 and j == 1:
                col.append(gui.Multiline("DITSC",
                                         size=(20, 4), key=f"({i},{j})", pad=(0, 0)))
                continue
            if i == 1 and j == 2:
                col.append(gui.Multiline("User Name(s) who would be\naccessing the ID",
                                         size=(20, 4), key=f"({i},{j})", pad=(0, 0)))
                continue
            if i == 3 and j == 1:
                col.append(gui.Multiline("Divisional CEO",
                                         size=(20, 4), key=f"({i},{j})", pad=(0, 0)))
                continue
            if i == 3 and j == 2:
                col.append(gui.Multiline("User Name(s) who would be\naccessing the ID",
                                         size=(20, 4), key=f"({i},{j})", pad=(0, 0)))
                continue
            if i == 4 and j == 2:
                col.append(gui.Multiline(
                    size=(20, 4), key=f"({i},{j})", pad=(0, 0)))
                col.append(
                    gui.Button('Exit', button_color=('white', 'red'), font=(
                               'Courier', 8, 'bold'))
                )

            col.append(gui.Multiline(
                size=(20, 4), key=f"({i},{j})", pad=(0, 0)))
        row.append(col)

    table = [header,
             [gui.Text('Network Login ID- Individual(AD Access)', size=(20, 4),
                       pad=(0, 0), justification='c'), row[0][0], row[0][1], row[0][2]],
             [gui.Text('Network Login ID- Shared/Generic(AD Access)', size=(20, 4),
                       pad=(0, 0), justification='c'), row[1][0], row[1][1], row[1][2]],
             [gui.Text('\nEmail ID- Individual', size=(20, 4), pad=(0, 0),
                       justification='c'), row[2][0], row[2][1], row[2][2]],
             [gui.Text('\nEmail ID- Shared/ Generic', size=(20, 4), pad=(0, 0),
                       justification='c'), row[3][0], row[3][1], row[3][2]],
             [gui.Text('Facility to reciecve mails from non-ITC E-mail ids(SMTP inbound)',
                       size=(20, 4), pad=(0, 0), justification='c'), row[4][0], row[4][1], row[4][2]]
             ]

    layout = [
        [gui.Text('Domain /Email /24*7 internet access',
                  font=('Courier', 16, 'bold'))],
    ]

    layout.append(table)

    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 8, 'bold')), gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                                                       font=('Courier', 10, 'bold'))]

    layout.append(button_row)

    window = gui.Window('Domain /Email /24*7 internet access', layout, font='Courier 10',
                        resizable=True, element_justification='c', keep_on_top=True)

    while True:
        event, values = window.read()
        if event in (gui.WIN_CLOSED, 'Exit'):
            break

        if event == 'Next':
            # print(values['(0,0)'])
            for i in range(5):
                for j in range(3):
                    key = f"_{i}{j}"
                    value = values[f'({i},{j})']
                    context.update({key: value})
            window.close()
            internetAccess2()

def internetAccess2():
    header = [
        gui.Text('IT Access required', size=(20, 2), justification='c'),
        gui.Text('Date till which access is required',
                 size=(20, 2), justification='c'),
        gui.Text('Additional approvals required to be enclosed',
                 size=(20, 2), justification='c'),
        gui.Text('Additional Information /Remarks',
                 size=(20, 2), justification='c')
    ]
    row = []
    for i in range(0, 5):
        col = []
        for j in range(0, 3):
            if i == 5-5 and j == 1:
                col.append(gui.Multiline("Head-Finance and MIS for non-\nmanagement employees",
                                         size=(20, 4), key=f"({i+5},{j})", pad=(0, 0)))
                continue
            if i == 5-5 and j == 2:
                col.append(gui.Multiline("Supervisor's Email Id",
                                         size=(20, 4), key=f"({i+5},{j})", pad=(0, 0)))
                continue
            if i == 6-5 and j == 1:
                col.append(gui.Multiline(
                    "For below level 6 and non\nmgmt. work justification and\napprovals to be provided\nseparately",
                    size=(20, 4), key=f"({i+5},{j})", pad=(0, 0)))
                continue
            if i == 7-5 and j == 1:
                col.append(gui.Multiline("NA", font=("Courier", 10, 'bold'),
                                         size=(20, 4), key=f"({i+5},{j})", pad=(0, 0)))
                continue
            if i == 8-5 and j == 1:
                col.append(gui.Multiline("Approval from Shared drive\nowner", font=("Courier", 10, 'bold'),
                                         size=(20, 4), key=f"({i+5},{j})", pad=(0, 0)))
                continue

            col.append(gui.Multiline(
                size=(20, 4), key=f"({i+5},{j})", pad=(0, 0)))
        row.append(col)

    table = [header,
           
             [gui.Text('Facility to send mails to non-ITC E-mail ids(SMTP outbound)',
                       pad=(0, 0), size=(20, 4), justification='c'), row[0][0], row[0][1], row[0][2]],
             [gui.Text('\n24*7 Internet Access', size=(20, 4), pad=(0, 0),
                       justification='c'), row[1][0], row[1][1], row[1][2]],
             [gui.Text('\nDefault AD Group to be Added ITD_All', size=(20, 4), pad=(
                 0, 0), justification='c'), row[2][0], row[2][1], row[2][2]],
             [gui.Text('Access to shared network drive <Shared drive name>', size=(
                 20, 4), pad=(0, 0), justification='c'), row[3][0], row[3][1], row[3][2]],
             [gui.Text('\nIT Asset Required Desktop /Laptop', size=(20, 4),
                       pad=(0, 0), justification='c'), row[4][0], row[4][1], row[4][2]]
             ]

    layout = [
        [gui.Text('Domain /Email /24*7 internet access',
                  font=('Courier', 16, 'bold'))],
        [gui.Text('part-2',
                  font=('Courier', 12, 'bold'))]
    ]

    layout.append(table)

    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 8, 'bold')), gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                                                       font=('Courier', 10, 'bold'))]

    layout.append(button_row)

    window = gui.Window('Domain /Email /24*7 internet access', layout, font='Courier 10',
                        resizable=True, element_justification='c')

    while True:
        event, values = window.read()
        if event in (gui.WIN_CLOSED, 'Exit'):
            break

        if event == 'Next':
            # print(values['(0,0)'])
            for i in range(5):
                for j in range(3):
                    key = f"_{i+5}{j}"
                    value = values[f'({i+5},{j})']
                    context.update({key: value})
            window.close()
            remote()
# end of the function internet access


def remote():
    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 10, 'bold')), gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                                                        font=('Courier', 10, 'bold'))]
    layout = [
        [gui.Text('Remote Access:-', font=('Courier', 16, 'bold'))],
        [gui.Text('Channel Name: VPN')],
        [gui.Text('Date till which access is required:', size=(30, 2)), gui.Input('Date', size=10, key='date_till_access1'),
            gui.CalendarButton('...', format="%d-%m-%y", close_when_date_chosen=True)],
        [gui.Text('Justification:'), gui.Push(), gui.Multiline(
            size=(20, 5), key='justification1')],
        [gui.Text('Approvals enclosed:'), gui.Push(),
         gui.Multiline(size=(20, 5), key='approval1')]

    ]

    layout.append(button_row)
    window = gui.Window('Remote Access', layout, font='Courier 10', keep_on_top=True,
                        resizable=True, element_justification='c')

    while True:
        event, values = window.read()
        if event == gui.WIN_CLOSED or event == 'Exit':
            break

        if event in ('Next'):

            context.update({
                'date_till_access1': values['date_till_access1'],
                'justification1': values['justification1'],
                'approval1': values['approval1']
            })
            window.close()
            Access_to_IT()

# end of the function remote access


def Access_to_IT():

    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 10, 'bold')), gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                                                        font=('Courier', 10, 'bold'))]
    layout = [
        [gui.Text('Access to IT Applications', font=('Courier', 16, 'bold'))],
        [gui.Text('Application Name:'), gui.Push(), gui.Multiline(
            size=(20, 4), key='application_name')],
        [gui.Text('ID Type:'), gui.Push(), gui.Combo(
            ['Individual', 'Shared', 'Generic'], key='id_type')],
        [gui.Text('Access profile or List of menu/ data options required:',
                  size=(20, 3)), gui.Push(), gui.Multiline(size=(20, 4), key='access_menu')],
        [gui.Text('Date till which access is required:', size=(20, 3)), gui.Push(), gui.Input('Date', size=(10, 1), key='date_till_access2'),
            gui.CalendarButton('...', format="%d-%m-%y", close_when_date_chosen=True)],
        [gui.Text('Justification:'), gui.Push(), gui.Multiline(
            size=(20, 4), key='justification2')],
        [gui.Text('Approvals enclosed:'), gui.Push(),
         gui.Multiline(
            "Approvals as\ndefined by\nthe\nApplication\nOwner for\nthe specific\napplication\nDITSC- for\nGeneric/\nShared user",
            size=(20, 4), key='approval2')],
        [gui.Text('Remarks:'), gui.Push(), gui.Multiline(
            size=(20, 4), key='remarks')]

    ]

    layout.append(button_row)
    window = gui.Window('Access to IT Applications', layout, keep_on_top=True,
                        font='Courier 10', resizable=True, element_justification='c')

    while True:
        event, values = window.read()
        if event == gui.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Next':
            #
            context.update({
                'application_name': values['application_name'],
                'id_type': values['id_type'],
                'access_menu': values['access_menu'],
                'date_till_access2': values['date_till_access2'],
                'justification2': values['justification2'],
                'approval2': values['approval2'],
                'remarks': values['remarks']
            })

            window.close()
            access_to_company()
# end of the function access to IT


def access_to_company():

    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 10, 'bold')), gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                                                        font=('Courier', 10, 'bold'))]
    layout = [
        [gui.Text('Access to company systems, applications, data and network through mobile devices', size=(
            30, 2), font=('Courier', 16, 'bold'))],
        [gui.Text('Company systems, applications, data details to which access is required: ', size=(
            20, 4)), gui.Push(), gui.Multiline(size=(20, 4), key='company_system')],
        [gui.Text('Name/ Model of the device: '), gui.Push(),
         gui.Multiline(size=(20, 4), key='model')],
        [gui.Text('IMEI/ PIN No: '), gui.Push(),
         gui.Multiline(size=(20, 4), key='imei')],

        [gui.Text('Date till which access is required:', size=(20, 3)), gui.Push(), gui.Input('Date', size=(10, 1), key='date_till_access3'),
            gui.CalendarButton('...', format="%d-%m-%y", close_when_date_chosen=True, key='date_till_access')],
        [gui.Text('Details of existing device in case the acces is being sought for a 2nd device (Maje/ IMEI): ',
                  size=(20, 6)), gui.Push(), gui.Multiline(size=(20, 6), key='details_of_exisiting_device')],
        [gui.Text('Additional approvals to be enclosed:', size=(20, 3)), gui.Push(), gui.Multiline(
            "1st time access:\nApproval from Divisional\nCEO manager designated\nby Divisional CEO> and CIO\n\nAccess on a 2nd device:\nApproval from DFC",
            size=(20, 4), key='additional_approval')]

    ]

    layout.append(button_row)
    window = gui.Window('Access to company systems, applications, data and network through mobile devices', layout, font='Courier 10',
                        resizable=True, element_justification='c', keep_on_top=True)

    while True:
        event, values = window.read()
        if event == gui.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Next':
            # print(values['Application_Name'])
            context.update({
                'company_system': values['company_system'],
                'model': values['model'],
                'imei': values['imei'],
                'date_till_access3': values['date_till_access3'],
                'additional_approval': values['additional_approval']
            })

            window.close()
            approval()

# end of the function access to company system


def approval():
    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 10, 'bold')), gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                                                        font=('Courier', 10, 'bold'))]
    layout = [
        [gui.Text('Approval of the Access ', font=('Courier',16,'bold'))],
        [gui.Text('Date:'), gui.Push(), gui.Input('Date', size=(10, 1), key='current_date'),
            gui.CalendarButton('...', format="%d-%m-%y", close_when_date_chosen=True, key='date_till_access')],
        [gui.Text("Superior's Name: "), gui.Push(),
         gui.Input(size=20, key='sup_name')],
        [gui.Text("HOD's /Unit Head's Name: "), gui.Push(),
         gui.Input(size=20, key='hod')],

        [gui.Text("\n")],
        [gui.Text('Approval from DMM /Unit IT Head ', font=('Courier',16,'bold'))],

        [gui.Text("Name: "), gui.Push(), gui.Input(size=20, key='head_name')],

    ]

    layout.append(button_row)
    window = gui.Window('Approvals', layout, keep_on_top=True,
                        font='Courier 10', resizable=True)

    while True:
        event, values = window.read()
        if event == gui.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Next':
            context.update({
                'sup_name': values['sup_name'],
                'hod': values['hod'],
                'current_date': values['current_date'],
                'head_name': values['head_name']
            })

            window.close()
            declaration()

# end of the function approval of the access


def declaration():
    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 10, 'bold')), gui.Push(), gui.Button('Submit', button_color=('black', 'sky blue'),
                                                        font=('Courier', 10, 'bold'))]

    layout = [
        [gui.Text('Declaration', font=('Courier',16,'bold'))],
        [gui.Text("Date:"), gui.Push(), gui.Input('Date', size=10, key='current_date'),
         gui.CalendarButton("..", format="%d-%m-%y", close_when_date_chosen=True)],
        [gui.Button('Click to read the Usage Policy Guidelines',
                    button_color=('black', 'grey'))],
        [gui.Checkbox(
            "I have read and understood the above mentioned\n Acceptable Usage Policy Guidelines", key='checked')],
        [gui.Text("Name of the User : {}".format(context['user_name']))],
        [gui.Text("Division/ SBU & Location:"), gui.Push(),
         gui.Input(size=20, key='division')],
        [gui.Text("Employee code:"), gui.Push(),
         gui.Input(size=20, key='emp_code')],
        # [gui.Text("\n")],
        button_row,
        [gui.Button('Click to Generate the pdf format of the Data Entered',
                    button_color=('black', 'grey'))]

    ]
    # layout.append(button_row)
    window = gui.Window("Declaration", layout,
                        # keep_on_top=True,
                        font='Courier 10', element_justification='c')

    submit = False
    while True:
        event, values = window.read()
        if event == gui.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Submit':
            submit = True

            if (values['checked'] == False):
                gui.popup("Accept the terms and condition")
            else:
                gui.popup("The data has been submitted successfully!")
                context.update({
                    'division': values['division'],
                    'emp_code': values['emp_code']
                })
        if event == 'Click to Generate the pdf format of the Data Entered':
            if submit is not True:
                gui.popup('Click on submit button to Submit the data!')
                continue
            else:
                render_to_doc()
                window.close()
        if event == 'Click to read the Usage Policy Guidelines':
            subprocess.Popen("Usage Policy Guidelines.pdf", shell=True)

# end of the function declaration


def render_to_doc():
    doc = DocxTemplate("userAccessForm.docx")

    doc.save(f"{context['first_name']}.docx")

    # create the copy of the file with the file name as user name
    copy_doc = DocxTemplate(
        f"{context['first_name']}.docx")
    # now render the context which is dictionary
    copy_doc.render(context)
    # save it again in the same directory
    copy_doc.save(f"{user_name}.docx")
    os.remove(f"{context['first_name']}.docx")

    # now convert to pdf
    # word_name=user_name+".dcox"
    # pdf_name=user_name+".pdf"
    docx2pdf.convert("{}.docx".format(user_name))
    # docx2pdf.convert(f"{context['user_name']}.docx","output.pdf")

    gui.popup("PDF file created successfully!")
    save_file()
    window.close()


# end of the function render_to_doc

def save_file():
    working_directory = os.getcwd()

    layout = [
        [gui.Text("Choose a path to save the pdf and word file",
                  font=('Courier', 10, 'bold'), size=(30, 2))],
    ]
    button_row = [gui.Button('Exit', button_color=('white', 'red'), font=(
        'Courier', 10, 'bold')), gui.Push(), gui.Button('Save')]

    word = f"{context['user_name']}.docx"
    pdf = f"{context['user_name']}.pdf"
    save = []
    save_path_memory = open("path_memory.txt", 'r')
    read = save_path_memory.read()
    save_path_memory.close()

    if (os.stat("path_memory.txt").st_size == 0):
        save = [gui.InputText(key='file_path'),
                gui.FolderBrowse(initial_folder=working_directory)],
    else:
        save = [gui.InputText(f"{read}", key='file_path'),
                gui.FolderBrowse(initial_folder=working_directory)]

    layout.append(save)
    layout.append(button_row)

    window = gui.Window("Save the File",layout)

    while True:
        event, values = window.read()
        
        try:
            save_path = values['file_path']
        except:
            pass

        if event in (gui.WIN_CLOSED, 'Exit'):
            break

        if event == "Save":
            save_path_memory = open("path_memory.txt", 'w')
            save_path_memory.write(save_path)
            save_path_memory.close()

            if not os.path.exists(save_path+"/"+f"{user_name}"):
                os.makedirs(save_path+"/"+f"{user_name}")

            # if not os.path.exists(save_path+"/"+"word_file"):
            #     os.makedirs(save_path+"/"+"word_file")

            word_path = values["file_path"] + f"/{user_name}/" + word
            shutil.move(word, word_path)

            pdf_path = values["file_path"] + f"/{user_name}/" + pdf
            shutil.move(pdf, pdf_path)

            gui.popup('File was saved successfully!')
            window.close()


# end of the function save_file()


# Main Programmm
layout = [
    [gui.Text('Demographic details', font=(
        'Courier', 16, 'bold'), justification='c')],
    [gui.Text('Date of request:', pad=(0, 0)), gui.Push(), gui.Input('Date', size=10, key='request_date'),
        gui.CalendarButton("..", pad=(0, 0), format="%d-%m-%y", close_when_date_chosen=True)],
    [gui.Text('Action:', pad=(0, 0)), gui.Push(), gui.Combo(
        ['Add', 'Change', 'Transfer'], key='action', size=18)],
    [gui.Text('First Name:', pad=(0, 0)), gui.Push(),
     gui.InputText(size=20, key='first_name')],
    [gui.Text('Last Name:', pad=(0, 0)), gui.Push(),
     gui.InputText(size=20, key='last_name')],
    [gui.Text('Unit Name:', pad=(0, 0)), gui.Push(),
     gui.InputText(size=20, key='unit_name')],
    [gui.Text('Designated and Grade/Level:', size=(20, 2), pad=(0, 0)),
     gui.Push(), gui.InputText(size=20, key='level')],
    [gui.Text('Type of access', pad=(0, 0)), gui.Push(), gui.Combo(
        ['Permanent', 'Temperory'], key='access_type', size=18)],
    [gui.Text('Employee Number:', pad=(0, 0)), gui.Push(),
     gui.InputText(size=20, key='emp_no')],
    [gui.Text("New Reporting Manager's name:", size=(20, 2), pad=(
        0, 0)), gui.Push(), gui.InputText(size=20, key='new_man')],
    [gui.Text("Previous Reporting Manager's name:", pad=(0, 0), size=(
        20, 2)), gui.Push(), gui.InputText(size=20, key='prev_man')],
    [gui.Text('Function/Department:', size=20, pad=(0, 0)),
     gui.Push(), gui.InputText(size=20, key='dept')],
    [gui.Text('Transfer From/To:', pad=(0, 0)), gui.Push(),
     gui.InputText(size=20, key='transfer_from_to')],
    [gui.Text('Validity Period:', pad=(0, 0)), gui.Push(),
     gui.Input(size=20, key='period')],

    [gui.Button('Exit', button_color=('white', 'red'), font=('Courier', 10, 'bold')),
        gui.Push(), gui.Button('Next', button_color=('black', 'sky blue'),
                               font=('Courier', 10, 'bold'))]
]


window = gui.Window('User Access Form', layout, font='Courier 10', keep_on_top=True,
                    resizable=True, element_justification='c')


while True:
    event, values = window.read()
    if event == gui.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Next':
        context.update({
            'request_date': values['request_date'],
            'action': values['action'],
            'first_name': values['first_name'],
            'last_name': values['last_name'],
            'level': values['level'],
            'access_type': values['access_type'],
            'emp_no': values['emp_no'],
            'new_man': values['new_man'],
            'prev_man': values['prev_man'],
            'dept': values['dept'],
            'transfer_from_to': values['transfer_from_to'],
            'period': values['period']
        })
        window.close()
        user_name = context['first_name']+" "+context['last_name']
        context.update({
            'user_name': user_name
        })
        internetAccess1()
