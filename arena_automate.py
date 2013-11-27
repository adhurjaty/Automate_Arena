from selenium import webdriver
import wx
import os
import sys
from itertools import izip, cycle
import base64
import re

'''
Params/Properties Dictionary keys:
part_number
revision
part_name
old_part_number
engineer
dco
dco_title
dco_desc
dco_number
'''

class Interface(wx.Frame):

    def __init__(self, parent, title, debug=False):
        
        super(Interface, self).__init__(parent, title=title,
                                        size=(350, 175))

        self.debug = debug
        self.login_panel = Login(self)
        self.login_panel.Hide()
        self.prompt_panel = PromptAction(self)
        self.prompt_panel.Hide()
        self.verify_panel = Verify(self)
        self.verify_panel.Hide()
        self.dco_panel = MakeDCO(self)
        self.dco_panel.Hide()

        self.dco = False #boolean indicating whether to add change to DCO

        ret_user = return_user()
        if not ret_user:
            self.SetTitle('Login')
            self.login_panel.Show()
        else:
            if debug:
                self.SetSizeWH(400, 70)
            else:
                self.SetSizeWH(295, 70)
            [self.email, self.password, self.engineer] = ret_user
            self.SetTitle('Select Action')
            self.prompt_panel.Show()

        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.login_panel, 1, wx.EXPAND)
        self.sizer.Add(self.prompt_panel, 1, wx.EXPAND)
        self.sizer.Add(self.verify_panel, 1, wx.EXPAND)
        self.sizer.Add(self.dco_panel, 1, wx.EXPAND)
        self.SetSizer(self.sizer)

        self.Centre()
        self.Show()

class Login(wx.Panel):
  
    def __init__(self, parent):
        super(Login, self).__init__(parent)

        self.parent = parent
            
        sizer = wx.GridBagSizer(5, 4)

        text = wx.StaticText(self, label="Wyatt Email:")
        sizer.Add(text, pos=(1, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.email = wx.TextCtrl(self)
        sizer.Add(self.email, pos=(1, 1), span=(1, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        text = wx.StaticText(self, label="Password:")
        sizer.Add(text, pos=(2, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.password = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        sizer.Add(self.password, pos=(2, 1), span=(1, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        buttonOk = wx.Button(self, label="&Ok", size=(90, 28))
        buttonClose = wx.Button(self, label="&Close", size=(90, 28))

        buttonOk.Bind(wx.EVT_BUTTON, self.login)
        buttonClose.Bind(wx.EVT_BUTTON, self.click_close)
        
        sizer.Add(buttonOk, pos=(4, 3))
        sizer.Add(buttonClose, pos=(4, 4), flag=wx.RIGHT|wx.BOTTOM, border=5)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(3)
        self.SetSizerAndFit(sizer)

    def click_close(self, e=None):
        self.parent.Close(True)

    def login(self, e=None):
        #get email and password and show error if one field is blank
        email = self.email.GetValue()
        password = self.password.GetValue()

        if not email or not password:
            show_error('Value Error', 'Please Enter Email and Password')

        #check that email and password are valid by putting them in Arena login
        #will exit with error if it fails. save the browser place in the parent
        self.parent.browser = enter_arena(email, password)

        #write email and password to file in My Documents
        with open('C:\\Users\\%s\\Documents\\login.dat'%os.environ.get('USERNAME'), 'w') as f:
            #encrypt password when writing the the login file
            f.write(email+'\n'+xor_crypt_string(password, encode=True))

        #lookup table corresponding from username to full name
        name_hash = {'adhurjaty': 'Anil Dhurjaty', 'rsleiman': 'Richard Sleiman',
                     'tkanusky': 'Thomas Kanusky', 'pneilson': 'Peter Neilson'}
        
        self.parent.login_panel.Hide()
        self.parent.SetTitle('Select Action')
        self.parent.SetSizeWH(295, 70)
        self.parent.RequestUserAttention() #notifies user that window needs attention
        self.parent.prompt_panel.Show() #switch to prompt action frame

class PromptAction(wx.Panel):
    def __init__(self, parent):
        super(PromptAction, self).__init__(parent)

        self.parent = parent

        if parent.debug:
            sizer = wx.GridBagSizer(1, 4)
        else:
            sizer = wx.GridBagSizer(1, 3)
        
        revise = wx.Button(self, label='Revise Part', size=(90, 28))
        revise.Bind(wx.EVT_BUTTON, self.revise_part)

        replace = wx.Button(self, label='Replace Part', size=(90, 28))
        replace.Bind(wx.EVT_BUTTON, self.replace_part)

        new = wx.Button(self, label='New Part', size=(90, 28))
        new.Bind(wx.EVT_BUTTON, self.new_part)

        sizer.AddMany([(revise, (0,0)), (replace, (0,1)),
                       (new, (0,2))])
        
        if parent.debug:
            test = wx.Button(self, label='Test', size=(90, 28))
            test.Bind(wx.EVT_BUTTON, self.test_button)

            sizer.Add(test, (0, 3))
        
        sizer.AddGrowableRow(0)
        self.SetSizerAndFit(sizer)
        

    def revise_part(self, e=None):
        self.parent.Hide()
        self.Hide()
        params = get_pdf(self)
        params.update(engineer=self.parent.engineer)
        self.parent.verify_panel.populate_form(**params)
        self.parent.SetTitle('Verify')
        self.parent.SetSizeWH(405, 330)
        self.parent.Show()
        self.parent.verify_panel.Show()

    def replace_part(self, e=None):
        
        self.parent.Hide()
        self.Hide()
        params = get_pdf(self)
        params.update(engineer=self.parent.engineer)
        self.parent.verify_panel.populate_form(2, **params)
        self.parent.SetTitle('Verify')
        self.parent.SetSizeWH(430, 400)
        self.parent.Show()
        self.parent.verify_panel.Show()

    def new_part(self, e=None):
        self.parent.Hide()
        self.Hide() 
        #get part_number, revision and path from user specifying drawing file
        params = get_pdf(self)

        #add to params dict to pass to create_part
        params.update(engineer=self.parent.engineer)

        self.parent.SetTitle('Verify')
        self.parent.verify_panel.populate_form(1, **params)
        self.parent.SetSizeWH(405, 365)
        self.parent.Show()
        self.parent.verify_panel.Show()

    def test_button(self, e=None):
        chromedriver = os.path.join(os.getcwd(), 'chrome\\chromedriver')
        os.environ['webdriver.chrome.driver'] = chromedriver
        br = webdriver.Chrome(chromedriver)
        br.get('http://www.reddit.com')
        links = [link.get_attribute('href') for link in br.find_elements_by_tag_name('a') if 'over' in link.text] 

        for link in links:
            br.get(link)
            a = raw_input('hit enter: ')

class Verify(wx.Panel):
    def __init__(self, parent):
        super(Verify, self).__init__(parent)

        self.parent = parent
        self.parent.dco = False

    def populate_form(self, new_part=0, **params):
        part_number = params['part_number']
        revision = params['revision']
        engineer = params['engineer']
        path = params['path']
        self.params = params
        self.new_part = new_part
        rows = 9 + new_part

        sizer = wx.GridBagSizer(rows, 4)

        self.pn_text = wx.TextCtrl(self)
        self.pn_text.WriteText(part_number)
        
        if new_part == 2:
            text = wx.StaticText(self, label="Replace:")
            sizer.Add(text, pos=(1, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

            self.rep_text = wx.TextCtrl(self)
            sizer.Add(self.rep_text, pos=(1, 1), span=(1, 4), 
                flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

            text = wx.StaticText(self, label="With:")
            sizer.Add(text, pos=(2, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

            sizer.Add(self.pn_text, pos=(2, 1), span=(1, 4), 
                flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        else:
            text = wx.StaticText(self, label="Part Number:")
            sizer.Add(text, pos=(1, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

            sizer.Add(self.pn_text, pos=(1, 1), span=(1, 4), 
                flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        text = wx.StaticText(self, label="Revision:")
        sizer.Add(text, pos=(rows-8, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.rev_text = wx.TextCtrl(self)
        self.rev_text.WriteText(revision)
        sizer.Add(self.rev_text, pos=(rows-8, 1), span=(1, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        if new_part:
            opt = ' (optional)' if new_part == 2 else ''
            text = wx.StaticText(self, label="Part Name%s:"%opt)
            sizer.Add(text, pos=(rows-7, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

            self.name_text = wx.TextCtrl(self)
            if 'part_name' in self.params.keys():
                self.name_text.WriteText(self.params['part_name'])
            sizer.Add(self.name_text, pos=(rows-7, 1), span=(1, 4), 
                flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        text = wx.StaticText(self, label="Engineer Name:")
        sizer.Add(text, pos=(rows-6, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.eng_text = wx.TextCtrl(self)
        self.eng_text.WriteText(engineer)
        sizer.Add(self.eng_text, pos=(rows-6, 1), span=(1, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        text = wx.StaticText(self, label="File:")
        sizer.Add(text, pos=(rows-5, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        #self.file_text = wx.TextCtrl(self)
        #self.file_text.WriteText(path)
        self.file_text = wx.StaticText(self, label=path.split('\\')[-1])
        sizer.Add(self.file_text, pos=(rows-5, 1), span=(1, 2), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        self.browse_button = wx.Button(self, label='Browse...', size = (75, 24))
        self.browse_button.Bind(wx.EVT_BUTTON, self.click_browse)
        sizer.Add(self.browse_button, pos=(rows-5, 3), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.image_text = wx.StaticText(self, label='Add Image (optional): ')
        sizer.Add(self.image_text, pos=(rows-4, 0), flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        self.image_button = wx.Button(self, label='Browse...', size = (75, 24))
        self.image_button.Bind(wx.EVT_BUTTON, self.image_browse)
        sizer.Add(self.image_button, pos=(rows-4, 1), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        if 'image' in self.params.keys():
            self.image_path = wx.StaticText(self, label=self.params['image'].split('\\')[-1])
            sizer.Add(self.image_path, pos=(rows-4, 2), span=(1, 2), 
                flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)
        

        cb = wx.CheckBox(self, label='Add to DCO')
        cb.Bind(wx.EVT_CHECKBOX, self.create_dco)
        sizer.Add(cb, pos=(rows-3,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)


        buttonOk = wx.Button(self, label="&Ok", size=(90, 28))
        buttonClose = wx.Button(self, label="&Close", size=(90, 28))

        buttonOk.Bind(wx.EVT_BUTTON, self.click_ok)
        buttonClose.Bind(wx.EVT_BUTTON, self.click_close)
        
        sizer.Add(buttonOk, pos=(rows-2, 3))
        sizer.Add(buttonClose, pos=(rows-2, 4), flag=wx.RIGHT|wx.BOTTOM, border=5)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(rows-3)

        self.SetSizerAndFit(sizer, True)

    def click_close(self, e=None):
        self.parent.Close(True)

    def click_ok(self, e=None):
        #self.parent.SetTitle('Select Action')
        self.parent.Hide()
        #self.parent.prompt_panel.Show()
        self.collect_values()

        if self.parent.dco:
            self.parent.Show()
            self.Hide()
            self.parent.SetSizeWH(380, 372)
            self.parent.SetTitle('Specify DCO')
            self.parent.dco_panel.Show()
        else:
            self.execute_action()

    def collect_values(self):
        self.params['part_number'] = self.pn_text.GetValue()
        self.params['revision'] = self.rev_text.GetValue()
        self.params['dco'] = self.parent.dco

        if self.new_part:
            self.params['part_name'] = self.name_text.GetValue()
        if self.new_part == 2:
            self.params['old_part_number'] = self.rep_text.GetValue()
        
    def click_browse(self, e=None):
        self.collect_values()
        properties = get_pdf(self)
        self.params.update(properties)
        self.populate_form(self.new_part, **self.params)

    def image_browse(self, e=None):
        self.collect_values()
        dlg = wx.FileDialog(self, 'Choose Part Image',
                            'M:\\Drawings\\Inventor part no 16xxxx',
                            '', 'JPEG or PNG (*.jpg *.png)|*.jpg;*.png;*.bmp')
    
        if dlg.ShowModal() == wx.ID_OK:
            filename = dlg.GetFilename()
            path = os.path.join(dlg.GetDirectory(), filename)
            path = path.replace('/','\\')
            self.params.update(image=path)
            self.populate_form(self.new_part, **self.params)

    #allows sequence to be called from MakeDCO class
    def execute_action(self):
        br = None
        try:
            br = self.parent.browser
        except:
            pass
        part_fnc = {0: update_part, 1: create_part, 2: replace_part}
        
        '''
        if we have already authenticated email and password by logging into Arena
        then the browser should be save and we do not have to log in again
        '''
        if br:
            br = part_fnc[self.new_part](br, **self.params)
        else:
            br = part_fnc[self.new_part](enter_arena(self.parent.email, self.parent.password), **self.params)

    def create_dco(self, e):
        if e.GetEventObject().GetValue():
            self.parent.dco = True
        else:
            self.parent.dco = False

class MakeDCO(wx.Panel):
    def __init__(self, parent):
        super(MakeDCO, self).__init__(parent)

        self.parent = parent

        sizer = wx.GridBagSizer(9, 4)

        text = wx.StaticText(self, label='Add to Existing DCO')
        sizer.Add(text, pos=(1,0), span=(1,4), flag=wx.ALIGN_CENTER)

        text = wx.StaticText(self, label="DCO Number:")
        sizer.Add(text, pos=(2, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.dco_text = wx.TextCtrl(self)
        sizer.Add(self.dco_text, pos=(2, 1), span=(1, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        text = wx.StaticText(self, label='Create New DCO')
        sizer.Add(text, pos=(3,0), span=(1,4), flag=wx.ALIGN_CENTER)

        text = wx.StaticText(self, label="DCO Title:")
        sizer.Add(text, pos=(4, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.title_text = wx.TextCtrl(self)
        sizer.Add(self.title_text, pos=(4, 1), span=(1, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        text = wx.StaticText(self, label="Description:")
        sizer.Add(text, pos=(5, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        self.desc_text = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        sizer.Add(self.desc_text, pos=(5, 1), span=(2, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        buttonOk = wx.Button(self, label="&Ok", size=(90, 28))
        buttonCancel = wx.Button(self, label="&Cancel", size=(90, 28))

        buttonOk.Bind(wx.EVT_BUTTON, self.click_ok)
        buttonCancel.Bind(wx.EVT_BUTTON, self.click_cancel)
        
        sizer.Add(buttonOk, pos=(8, 3))
        sizer.Add(buttonCancel, pos=(8, 4), flag=wx.RIGHT|wx.BOTTOM, border=5)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(6)

        self.SetSizerAndFit(sizer)

    def click_ok(self, e=None):
        
        dco_num = self.dco_text.GetValue()
        dco_title = self.title_text.GetValue()
        dco_desc = self.desc_text.GetValue()
        
        if dco_num:
            self.parent.verify_panel.params['dco_number'] = dco_num
        elif dco_title:
            self.parent.verify_panel.params['dco_title'] = dco_title
            self.parent.verify_panel.params['dco_desc'] = dco_desc
        else:
            self.parent.dco = False

        self.parent.Hide()
        self.parent.verify_panel.execute_action()

    def click_cancel(self, e=None):
        self.parent.dco = False
        self.Hide()
        self.parent.SetTitle('Select Action')
        self.parent.verify_panel.Show()

def return_user():
    try:
        with open('C:\\Users\\%s\\Documents\\login.dat'%os.environ.get('USERNAME'), 'r') as f:
            [email, password] = f.read().split('\n')
            password = xor_crypt_string(password, decode=True)
            
            name_hash = {'adhurjaty': 'Anil Dhurjaty', 'rsleiman': 'Richard Sleiman',
                     'tkanusky': 'Thomas Kanusky', 'pneilson': 'Peter Neilson'}
            engineer = name_hash[email.split('@')[0]]
            
            return [email, password, engineer]
    except:
        return False
    
def get_pdf(parent):
    
    dlg = wx.FileDialog(parent, 'Choose Part Drawing',
                            'M:\\Drawings\\Inventor part no 16xxxx',
                            '', '*.pdf')
    
    if dlg.ShowModal() == wx.ID_OK:
        filename = dlg.GetFilename()
        path = os.path.join(dlg.GetDirectory(), filename)
        path = path.replace('/','\\')

        filename = filename.split('-')
        part_number = filename[0] #first part of filename is part number
        revision = filename[-1].split('.')[0] #last is revision, remove '.pdf'
    
        options = False
        if len(filename) == 3: #if the part has -XX options, set option flag to true
            options = True
        
        return dict(part_number=part_number, revision=revision, path=path, options=options)
    show_error('File Error', 'You must select PDF file.')

def enter_arena(email, password):
    chromedriver = os.path.join(os.getcwd(), 'chrome\\chromedriver')
    os.environ['webdriver.chrome.driver'] = chromedriver
    br = webdriver.Chrome(chromedriver)
    br.get('https://app.bom.com/items/list-main') #go to items page
                                                  #will be redirected to login
    #fill in form and submit
    br.find_element_by_name('email').send_keys(email)
    br.find_element_by_name('password').send_keys(password)
    br.find_element_by_name('password').submit()

    error_info = br.find_elements_by_id('loginErrorInfo')
    if error_info:
        msg = error_info[0].find_element_by_tag_name('li').text
        br.quit()
        show_error('Login Error', msg)
    
    return br

def create_part(br, **properties):
    part_number = properties['part_number']
    part_name = properties['part_name']
    revision = properties['revision']
    engineer = properties['engineer']
    path = properties['path']
    
    #items page
    br.get(br.find_element_by_link_text('New Item').get_attribute('href'))
    
    #input part information page
    form = br.find_element_by_name('DataEntryForm')
    
    click_in_list(form.find_element_by_name('form_category_id'), 'option', 'Part')
    
            
    #search through elements whose name is 'format_field_values'
    for item in form.find_elements_by_name('format_field_values'): 
        if item.is_displayed: #and if the field is diplayed, enter part number in field
            item.send_keys(part_number)
            break

    form.find_element_by_name('form_version').send_keys(revision)
    form.find_element_by_name('form_item_name').send_keys(part_name)
    
    click_in_list(form.find_element_by_name('form_engineer'), 'option', engineer)
    
    form.find_elements_by_name('form_off_the_shelf_p')[1].click() #click Made-to-Specification
    form.submit()

    #check for errors
    edit_errors = br.find_elements_by_id('EditError')
    if edit_errors:
        msg = edit_errors[0].find_element_by_tag_name('li').text
        br.quit()
        show_error('Part Exists Error', msg, br)

    #part specs page
    #<add image>
    
    if 'image' in properties.keys():
        #click_in_list(br.find_element_by_id('SpecHeaderUploadImageLink'), 'td', 'Select Image')
        for option in br.find_element_by_id('SpecHeaderUploadImageLink').find_elements_by_tag_name('td'):
            if option.get_attribute('class') == 'TDViewBtn' and 'Select Image' in option.text:
                option.click()
                break
                
        #br.find_element_by_name('form_image_file_name').clear()
        br.find_element_by_name('form_image_file_name').send_keys(properties['image'])
    
    #</add image>
    
    br = go_to_tab(br, 'Files')

    #files page
    click_in_list(br.find_element_by_id('AttachHeaderCommands'), 'td', 'Add New Files')
    
    #add files page
    form = br.find_element_by_id('MultiPartAction_DataEntryForm')
    #add_file(form, path, part_number, revision, engineer)
    
    form.find_element_by_name('file_to_upload_0').send_keys(path)
    try:
        table = form.find_element_by_id('mfu_sm_0_0_file_info')
    except:
        table = form
    while not table.find_elements_by_name('form_file_identifier'):
        form.find_element_by_name('file_to_upload_0').send_keys(path)
    
    table.find_element_by_name('form_file_identifier').clear()
    table.find_element_by_name('form_file_identifier').send_keys(part_number)
    table.find_element_by_name('form_edition_identifier').send_keys(revision)

    click_in_list(table.find_element_by_name('form_file_author'), 'option', engineer)
    
    form.find_element_by_name('submitFileForm').click()

    check_dco(br, **properties)
            
    completed('Item Created', 'Successfully Created Item', br)
    
def update_part(br, **properties):
    part_number = properties['part_number']
    revision = properties['revision']
    engineer = properties['engineer']
    path = properties['path']
    
    #items page
    br = search_item(br, part_number)

    #part page
    br = working_rev(br) #go to working revision in drop-down
    br = go_to_tab(br, 'Specs')
    
    #<add image>
    if 'image' in properties.keys():
        #click_in_list(br.find_element_by_id('SpecHeaderUploadImageLink'), 'td', 'Select Image')
        for option in br.find_element_by_id('SpecHeaderUploadImageLink').find_elements_by_tag_name('td'):
            if option.get_attribute('class') == 'TDViewBtn' and 'Select Image' in option.text:
                option.click()
                break
                
        #br.find_element_by_name('form_image_file_name').clear()
        br.find_element_by_name('form_image_file_name').send_keys(properties['image'])
    
    #</add image>
    
    br = go_to_tab(br, 'Files')

    #files page
    
    br.get(br.find_element_by_link_text('Update').get_attribute('href'))

    #update file page
    form = br.find_element_by_id('MultiPartAction_DataEntryForm')
    ops = form.find_elements_by_name('form_storage_method')
    for o in ops:
        if o.get_attribute('value') == '0':
            o.click()
            break

    add_file(form, path, part_number, revision, engineer)

    check_dco(br, **properties)
        
    completed('Item Revised', 'Successfully Revised Item', br)

def replace_part(br, **properties):
    part_number = properties['old_part_number']
    revision = properties['revision']
    engineer = properties['engineer']
    path = properties['path']

    #<obsolete item>
    br = search_item(br, part_number)

    'get revision letter from soon-to-be-deprecated part'
    old_rev = next((re.search('[0-9A-Z]', o.text).group(0)
                    for o in br.find_element_by_id('screen-text').find_elements_by_tag_name('option')
                    if 'Working' not in o.text))
    
    'go to working revision'
    br = working_rev(br)
    'go to specs tab'
    br = go_to_tab(br, 'Specs')

    if not properties['part_name']:
        div = br.find_element_by_id('object-header')
        properties['part_name'] = div.find_element_by_tag_name('h2').text


    click_in_list(br.find_element_by_id('SpecHeaderEditItemLink'), 'td', 'Edit Information')

    form = br.find_element_by_name('DataEntryForm')
    item_number_fields = form.find_elements_by_name('format_field_values')
    for item in item_number_fields: #search through elements whose name is 'format_field_values'
        if item.is_displayed: #and if the field is diplayed, enter part number in field
            item.clear()
            item.send_keys('X'+part_number)
            break

    #make sure all required drop-down menus have values to prevent errors
    div = form.find_element_by_id('attsCatDriven')
    for sel in div.find_elements_by_tag_name('select'):
        click_in_list(sel, 'option', 'No')

    form.find_element_by_name('submit').click()

    if br.find_elements_by_id('EditError'):
        show_error('Existing part error',
                   br.find_element_by_id('EditError').find_element_by_tag_name('li').text)
    #</obsolete item>

    if properties['dco']:
        properties['revision'] = old_rev
        if 'dco_number' in properties.keys():
            br = open_obsolete_dco(br, **properties)
        elif 'dco_title' in properties.keys():
            [br, dco_number] = new_obsolete_dco(br, **properties)
            del properties['dco_title']
            properties['dco_number'] = dco_number
        properties['revision'] = revision

    create_part(br, **properties)

def check_dco(br, **properties):
    if properties['dco']:
        if 'dco_number' in properties.keys():
            add_to_dco(br, **properties)
        elif 'dco_title' in properties.keys():
            create_dco(br, **properties)
            
def create_dco(br, **properties):

    br = go_to_actions_dco(br, **properties)

    #select actions page
    table = br.find_element_by_id('MultiPartAction_DataEntryForm')
    for cb in table.find_elements_by_tag_name('input'):
        if 'Release to Production' in cb.find_element_by_xpath('../..').text\
           and cb.get_attribute('type') != 'hidden':
            cb.click()
    
    br.find_element_by_name('submitForm').click()

    finish_dco(br, **properties)

def add_to_dco(br, **properties):
    dco_number = properties['dco_number']
    revision = properties['revision']
    engineer = properties['engineer']

    br.get(br.find_element_by_id('revBarCommandBtnContainer').
           find_element_by_link_text('Add to Change').get_attribute('href'))
    table = br.find_element_by_id('topTTable')

    found_dco = False
    for links in table.find_elements_by_tag_name('a'):
        if dco_number in links.text:
            links.find_element_by_xpath('../..').find_element_by_name('top_checkbox').click()
            found_dco = True
            break

    if not found_dco:
        show_error('Invalid DCO', 'DCO number is invalid')
    
    br.find_element_by_id('MultiPartAction_FindChange').find_element_by_name('submitForm').click()

    #select actions page
    table = br.find_element_by_id('MultiPartAction_DataEntryForm')
    for cb in table.find_elements_by_tag_name('input'):
        if 'Release to Production' in cb.find_element_by_xpath('../..').text\
           and cb.get_attribute('type') != 'hidden':
            cb.click()
            
    br.find_element_by_name('submitForm').click()

    finish_dco(br, **properties)

def new_obsolete_dco(br, **properties):

    br = go_to_actions_dco(br, **properties)

    #select actions page
    table = br.find_element_by_id('MultiPartAction_DataEntryForm')

    for cb in table.find_elements_by_tag_name('input'):
        if 'Deprecate Item' in cb.find_element_by_xpath('../..').text\
           and cb.get_attribute('type') != 'hidden':
            cb.click()

    br.find_element_by_name('submitForm').click()
    
    br = finish_dco(br, **properties)

    dco_number = br.find_element_by_id('ObjectHeader').find_element_by_tag_name('h4').text.split('#')[-1]
    dco_number = dco_number.split(' ')[0]
    
    #go back to items page
    br.get(br.find_element_by_link_text('Production Workspace Items').get_attribute('href'))

    return br, dco_number

def open_obsolete_dco(br, **properties):
    dco_number = properties['dco_number']
    revision = properties['revision']
    engineer = properties['engineer']

    br.get(br.find_element_by_id('revBarCommandBtnContainer').
           find_element_by_link_text('Add to Change').get_attribute('href'))
    table = br.find_element_by_id('topTTable')

    found_dco = False
    for links in table.find_elements_by_tag_name('a'):
        if dco_number in links.text:
            links.find_element_by_xpath('../..').find_element_by_name('top_checkbox').click()
            found_dco = True
            break

    if not found_dco:
        show_error('Invalid DCO', 'DCO number is invalid')
    
    br.find_element_by_id('MultiPartAction_FindChange').find_element_by_name('submitForm').click()

    #select actions page
    table = br.find_element_by_id('MultiPartAction_DataEntryForm')

    for cb in table.find_elements_by_tag_name('input'):
        if 'Deprecate Item' in cb.find_element_by_xpath('../..').text\
           and cb.get_attribute('type') != 'hidden':
            cb.click()

    br.find_element_by_name('submitForm').click()
    
    br = finish_dco(br, **properties)

    #go back to items page
    br.get(br.find_element_by_link_text('Production Workspace Items').get_attribute('href'))

    return br

def go_to_actions_dco(br, **properties):
    title = properties['dco_title']
    description = properties['dco_desc']
    revision = properties['revision']
    engineer = properties['engineer']

    #select Add to Change button on parts page
    br.get(br.find_element_by_id('revBarCommandBtnContainer').
           find_element_by_link_text('Add to Change').get_attribute('href'))
    br.find_element_by_id('MultiPartAction_CreateNewChange').\
        find_element_by_name('submitForm').click()

    form = br.find_element_by_id('MultiPartAction_DataEntryForm')
    form.find_element_by_id('DCO-(# auto generated)').click()
    form.find_element_by_name('form_change_title').send_keys(title)
    form.find_element_by_id('formChangeDesc').send_keys(description)

    for opt in form.find_element_by_name('form_routings').find_elements_by_tag_name('option'):
        if 'Hardware - No Change to Published Specs' in opt.text:
            opt.click()
            break

    form.find_element_by_name('form_custom_45976').send_keys(engineer)
    form.find_element_by_name('submit').click()

    #specify items page
    while True:
        try:
            br.find_element_by_name('submitForm').click()
            break
        except:
            pass
    

    return br

def finish_dco(br, **properties):
    revision = properties['revision']
    engineer = properties['engineer']
    
    #enter revision page
    form = br.find_element_by_id('MultiPartAction_DataEntryForm')

    for val in form.find_elements_by_tag_name('input'):
        if 'form_version_num_' in val.get_attribute('name'):
            val.send_keys(revision)
            break

    for box in form.find_elements_by_tag_name('input'):
        if box.get_attribute('name') and 'form_version_views_'\
           in box.get_attribute('name') and not box.is_selected()\
           and box.get_attribute('type') != 'hidden':
            box.click()

    br.find_element_by_name('submitForm').click()

    if br.find_elements_by_id('EditError'):
        show_error('DCO Error', br.find_element_by_id('EditError').find_element_by_tag_name('li').text)

    return br

def go_to_tab(br, tab):
    lst = br.find_element_by_id('views')
    try:
        br.get(lst.find_element_by_link_text(tab).get_attribute('href'))
    except:
        pass
    return br

def add_file(form, path, part_number, revision, engineer):
    form.find_element_by_name('file_to_upload_0').send_keys(path)
    try:
        table = form.find_element_by_id('mfu_sm_0_0_file_info')
    except:
        table = form
    while not table.find_elements_by_name('form_file_identifier'):
        pass
    table.find_element_by_name('form_file_identifier').clear()
    table.find_element_by_name('form_file_identifier').send_keys(part_number)
    table.find_element_by_name('form_edition_identifier').send_keys(revision)
    authors = table.find_element_by_name('form_file_author')
    for a in authors.find_elements_by_tag_name('option'):
        if a.text == engineer:
            a.click()
            break
    form.find_element_by_name('submitFileForm').click()
    return form

def search_item(br, pn):
    br.find_element_by_name('search_textfield').send_keys(pn)
    br.find_element_by_name('SearchGo').click()

    error_info = br.find_elements_by_id('PbopHeader')
    if error_info:
        show_error('Non matching error', 'No item matches that part number')

    if 'list-main' in br.current_url.split('/'): #if return search results
        for link in br.find_elements_by_tag_name('a'):
            if pn in link.text:
                br.get(link.get_attribute('href')) #go to first item in list
                break

    return br

def working_rev(br):
    revs = br.find_element_by_name('display_revision')
    for r in revs.find_elements_by_tag_name('option'):
        if 'Working Revision' in r.text:
            r.click() #select working revision for part
            break
    return br

def click_in_list(form, tag, text):
    for option in form.find_elements_by_tag_name(tag):
        if text in option.text:
            option.click()
            break
    
def show_error(title, msg, browser=None):
    if browser:
        browser.quit()
    wx.MessageBox(msg, title, wx.OK|wx.ICON_ERROR)
    app.Destroy()
    sys.exit(1)

def completed(title, msg, browser=None):
    if browser:
        browser.quit()
    wx.MessageBox(msg, title, wx.OK|wx.ICON_INFORMATION)
    app.Destroy()
    sys.exit(0)


def xor_crypt_string(data, key='adhurjaty', encode=False, decode=False):
    if decode:
        data = base64.decodestring(data)
    xored = ''.join(chr(ord(x) ^ ord(y)) for (x,y) in izip(data, cycle(key)))
    if encode:
        return base64.encodestring(xored).strip()
    return xored

if __name__ == '__main__':

    app = wx.App()
    
    if len(sys.argv) > 1 and sys.argv[1] == '-debug':
        Interface(None, title='Login', debug=True)
    else:
        Interface(None, title='Login')
    app.MainLoop()
