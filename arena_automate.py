import Tkinter as Tk
from selenium import webdriver
import sys
import tkFileDialog
import tkMessageBox
import tkSimpleDialog
import os
from itertools import izip, cycle
import base64

class MyApp(Tk.Tk):
    def __init__(self):
        Tk.Tk.__init__(self)

        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = Tk.Frame(self)
        container.pack(side='top', fill='both', expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        #construct frames hash where separate windows are stored
        self.frames = {}
        for F in (PromptAction, Login):
            self.frames[F] = F(container, self)
            self.frames[F].grid(row=0, column=0, sticky='nsew')
            #creates layout such that frames are ontop of one another

        ret_user = return_user()
        if ret_user:
            [self.email, self.password, self.engineer] = ret_user
            self.show_frame(PromptAction)
        else:
            self.show_frame(Login)

    def show_frame(self, c):
        self.frames[c].tkraise()
        self.frames[c].focus_force()

class Login(Tk.Frame):
    def __init__(self, parent, controller):
        Tk.Frame.__init__(self, parent)
        self.parent = parent
        self.controller = controller
        
        controller.geometry('300x160')
        controller.title('Enter your information')
        
        #entrys with not shown text
        self.user = self.make_entry("Email Address Used in Arena:", 16)
        self.password = self.make_entry("Password:", 16, show="*")
        
        #button to attempt to login
        b = Tk.Button(self, borderwidth=4, text="Login", width=10, pady=8, command=self.login)
        b.pack(side=Tk.BOTTOM)
        self.password.bind('<Return>', self.login)

    def login(self, event=None):
        #get email and password and show error if one field is blank
        email = self.user.get()
        password = self.password.get()
        if not email or not password:
            self.controller.destroy()
            show_error('Value Error', 'Please Enter Email and Password')

        #check that email and password are valid by putting them in Arena login
        #will exit with error if it fails. save the browser place in the parent
        self.controller.browser = enter_arena(email, password)

        #write email and password to file in My Documents
        with open('C:\\Users\\%s\\Documents\\login.dat'%os.environ.get('USERNAME'), 'w') as f:
            #encrypt password when writing the the login file
            f.write(email+'\n'+xor_crypt_string(password, encode=True))

        #lookup table corresponding from username to full name
        name_hash = {'adhurjaty': 'Anil Dhurjaty', 'rsleiman': 'Richard Sleiman',
                     'tkanusky': 'Thomas Kanusky', 'pneilson': 'Peter Neilson'}

        #look up engineer from hash after removing email domain
        engineer = name_hash[email.split('@')[0]]

        #store these variables in the parent for usage in the PromptAction class
        self.controller.email = email
        self.controller.password = password
        self.controller.engineer = engineer

        #change to PromptAction class frame
        self.controller.show_frame(PromptAction)

    #creates entry field in window
    def make_entry(self, caption, width=None, **options):
        Tk.Label(self, text=caption).pack(side=Tk.TOP)
        entry = Tk.Entry(self, **options)
        if width:
            entry.config(width=width)
            entry.pack(side=Tk.TOP, padx=10, fill=Tk.BOTH)
        return entry

class PromptAction(Tk.Frame):
    #window where user selects to add new part, revise a part, or replace a part
    
    def __init__(self, parent, controller):
        Tk.Frame.__init__(self, parent)
        
        #------ constants for controlling layout ------
        button_width = 6      ### (1)
        
        button_padx = "2m"    ### (2)
        button_pady = "1m"    ### (2)

        buttons_frame_padx =  "3m"   ### (3)
        buttons_frame_pady =  "2m"   ### (3)		
        buttons_frame_ipadx = "3m"   ### (3)
        buttons_frame_ipady = "1m"   ### (3)
        # -------------- end constants ----------------
        self.controller = controller
        self.controller.title('Select Action')
        self.controller.geometry('150x50')   

        button1 = Tk.Button(self, command=self.revise, padx=buttons_frame_padx)
        button1.configure(text="Revise Part")
        button1.focus_force()       
        button1.configure( 
                width=button_width,  ### (1)
                padx=button_padx,    ### (2) 
                pady=button_pady     ### (2)
                )

        button1.pack(side=Tk.LEFT)	
        button1.bind("<Return>", self.revise)  
        
        button2 = Tk.Button(self, command=self.replace)
        button2.configure(text="Replace Part")  
        button2.configure( 
                width=button_width,  ### (1)
                padx=button_padx,    ### (2) 
                pady=button_pady     ### (2)
                )

        button2.pack(side=Tk.LEFT)
        button2.bind("<Return>", self.replace)

        button3 = Tk.Button(self, command=self.new_part)
        button3.configure(text="New Part")  
        button3.configure( 
                width=button_width,  ### (1)
                padx=button_padx,    ### (2) 
                pady=button_pady     ### (2)
                )

        button3.pack(side=Tk.LEFT)
        button3.bind("<Return>", self.new_part)
            
    def revise(self):
        #hide the main window
        self.controller.withdraw()
        #get part_number, revision and path from user specifying drawing file
        params = get_pdf()
        params.update(engineer=self.controller.engineer)
        self.controller.destroy()

        '''
        if we have already authenticated email and password by logging into Arena
        then the browser should be save and we do not have to log in again
        '''
        try:
            update_part(self.controller.browser, **params)
        except:
            update_part(enter_arena(self.controller.email, self.controller.password), **params)
    
    def replace(self): 
        self.controller.destroy()
            
    def new_part(self):  
        self.controller.withdraw()
        #get part_number, revision and path from user specifying drawing file
        params = get_pdf()
        #get the part name from user directly
        part_name = tkSimpleDialog.askstring('Part Name', 'Enter Part Name')

        #add to params dict to pass to create_part
        params.update(engineer=self.controller.engineer)
        params.update(part_name=part_name)
        self.controller.destroy()

        '''
        if we have already authenticated email and password by logging into Arena
        then the browser should be save and we do not have to log in again
        '''
        try:
            create_part(self.controller.browser, **params)
        except:
            create_part(enter_arena(self.controller.email, self.controller.password), **params)

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
    
def get_pdf():
    master = Tk.Tk()
    master.withdraw() #hide dummy window
    #ask user to find pdf file
    path = tkFileDialog.askopenfilename(title='Choose PDF Drawing',
                                     filetypes=[('PDF Files', '.pdf')],
                                     initialdir='M:\Drawings\Inventor part no 16xxxx')
    path = path.replace('/','\\')
    master.destroy()
    if path:
        filename = os.path.basename(path).split('-')
        part_number = filename[0] #first part of filename is part number
        revision = filename[-1].split('.')[0] #last is revision, remove '.pdf'
    
        options = False
        if len(filename) > 3: #if the part has -XX options, set option flag to true
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
    new_part_link = br.find_element_by_link_text('New Item').get_attribute('href')
    br.get(new_part_link)
    
    #input part information page
    form = br.find_element_by_name('DataEntryForm')
    category = form.find_element_by_name('form_category_id')
    for option in category.find_elements_by_tag_name('option'):
        if option.text == 'Part':
            option.click()
            
    item_number_fields = form.find_elements_by_name('format_field_values')
    for item in item_number_fields: #search through elements whose name is 'format_field_values'
        if item.is_displayed: #and if the field is diplayed, enter part number in field
            item.send_keys(part_number)
            break

    form.find_element_by_name('form_version').send_keys(revision)
    form.find_element_by_name('form_item_name').send_keys(part_name)
    
    engineers = form.find_element_by_name('form_engineer')
    for e in engineers.find_elements_by_tag_name('option'):
        if e.text == engineer:
            e.click()
    form.find_elements_by_name('form_off_the_shelf_p')[1].click() #click Made-to-Specification
    form.submit()

    #check for errors
    edit_errors = br.find_elements_by_id('EditError')
    if edit_errors:
        msg = edit_errors[0].find_element_by_tag_name('li').text
        br.quit()
        show_error('Part Exists Error', msg)

    #part specs page
    br = go_to_files(br)

    #files page
    add_file = br.find_element_by_id('AttachHeaderCommands')
    for val in add_file.find_elements_by_tag_name('td'):
        if val.text == 'Add New Files':
            val.click()
            break
    
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
    authors = table.find_element_by_name('form_file_author')
    for a in authors.find_elements_by_tag_name('option'):
        if a.text == engineer:
            a.click()
            break
    form.find_element_by_name('submitFileForm').click()
    completed('Item Created', 'Successfully Created Item', br)
    
def update_part(br, **properties):
    part_number = properties['part_number']
    revision = properties['revision']
    engineer = properties['engineer']
    path = properties['path']
    
    #items page
    br.find_element_by_name('search_textfield').send_keys(part_number)
    br.find_element_by_name('SearchGo').click()

    error_info = br.find_elements_by_id('PbopHeader')
    if error_info:
        show_error('Non matching error', 'No item matches that part number')
    
    if 'list-main' in br.current_url.split('/'): #if return search results
        part_link = br.find_element_by_link_text(part_number).get_attribute('href')
        br.get(part_link) #go to first item in list

    #part page
    br = go_to_files(br)

    #files page
    revs = br.find_element_by_name('display_revision')
    for r in revs.find_elements_by_tag_name('option'):
        if r.text == 'Working Revision':
            r.click() #select working revision for part
            break
    
    br.get(br.find_element_by_link_text('Update').get_attribute('href'))

    #update file page
    form = br.find_element_by_id('MultiPartAction_DataEntryForm')
    ops = form.find_elements_by_name('form_storage_method')
    for o in ops:
        if o.get_attribute('value') == '0':
            o.click()

    add_file(form, path, part_number, revision, engineer)
    completed('Item Revised', 'Successfully Revised Item', br)

def completed(title, msg, browser):
    master = Tk.Tk()
    master.withdraw()
    browser.quit()
    tkMessageBox.showinfo(title, msg)
    master.destroy()
    
def show_error(title, msg):
    master = Tk.Tk()
    master.withdraw()
    try:
        myapp.destroy()
    except:
        pass
    tkMessageBox.showerror(title, msg)
    master.destroy()
    sys.exit(1)
    
def go_to_files(br):
    br.get(br.find_elements_by_link_text('Files')[1].get_attribute('href'))
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

def xor_crypt_string(data, key='adhurjaty', encode=False, decode=False):
    if decode:
        data = base64.decodestring(data)
    xored = ''.join(chr(ord(x) ^ ord(y)) for (x,y) in izip(data, cycle(key)))
    if encode:
        return base64.encodestring(xored).strip()
    return xored

def automate():
    create_part(enter_arena('adhurjaty@wyatt.com', 'Wyattme23'),
                path='C:/Users/adhurjaty/Desktop/123456-A.pdf',
                part_number='123456', revision='A', part_name='test',
                engineer='Anil Dhurjaty')

#automate()


myapp = MyApp()
myapp.mainloop()



