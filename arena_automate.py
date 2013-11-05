import Tkinter as Tk
from selenium import webdriver
import sys
import tkFileDialog
import tkMessageBox
import tkSimpleDialog
import os
from itertools import izip, cycle
import base64

class MyApp:
    def __init__(self, parent):
		
        #------ constants for controlling layout ------
        button_width = 6      ### (1)
        
        button_padx = "2m"    ### (2)
        button_pady = "1m"    ### (2)

        buttons_frame_padx =  "3m"   ### (3)
        buttons_frame_pady =  "2m"   ### (3)		
        buttons_frame_ipadx = "3m"   ### (3)
        buttons_frame_ipady = "1m"   ### (3)
        # -------------- end constants ----------------
        
        self.myParent = parent   
        self.buttons_frame = Tk.Frame(parent)
        
        self.buttons_frame.pack(    ### (4)
            ipadx=buttons_frame_ipadx,  ### (3)
            ipady=buttons_frame_ipady,  ### (3)
            padx=buttons_frame_padx,    ### (3)
            pady=buttons_frame_pady,    ### (3)
            )    
        

        self.button1 = Tk.Button(self.buttons_frame, command=self.revise)
        self.button1.configure(text="Revise Part")
        self.button1.focus_force()       
        self.button1.configure( 
                width=button_width,  ### (1)
                padx=button_padx,    ### (2) 
                pady=button_pady     ### (2)
                )

        self.button1.pack(side=Tk.LEFT)	
        self.button1.bind("<Return>", self.revise)  
        
        self.button2 = Tk.Button(self.buttons_frame, command=self.replace)
        self.button2.configure(text="Replace Part")  
        self.button2.configure( 
                width=button_width,  ### (1)
                padx=button_padx,    ### (2) 
                pady=button_pady     ### (2)
                )

        self.button2.pack(side=Tk.LEFT)
        self.button2.bind("<Return>", self.replace)

        self.button3 = Tk.Button(self.buttons_frame, command=self.new_part)
        self.button3.configure(text="New Part")  
        self.button3.configure( 
                width=button_width,  ### (1)
                padx=button_padx,    ### (2) 
                pady=button_pady     ### (2)
                )

        self.button3.pack(side=Tk.LEFT)
        self.button3.bind("<Return>", self.new_part) 
            
    def revise(self):
        self.myParent.withdraw()
        [email, password, engineer] = authenticate()
        params = get_pdf()
        params.update(engineer=engineer)
        self.myParent.destroy()
        update_part(enter_arena(email, password), **params)
    
    def replace(self): 
        self.myParent.destroy()
            
    def new_part(self):  
        self.myParent.withdraw()
        [email, password, engineer] = authenticate()
        params = get_pdf()
        part_name = tkSimpleDialog.askstring('Part Name', 'Enter Part Name')
        params.update(engineer=engineer)
        params.update(part_name=part_name)
        self.myParent.destroy()
        create_part(enter_arena(email, password), **params)

def authenticate():
    try:
        with open('C:\\Users\\%s\\Documents\\login.dat'%os.environ.get('USERNAME'), 'r') as f:
            [email, password] = f.read().split('\n')
            password = xor_crypt_string(password, decode=True)
    except:
        email = tkSimpleDialog.askstring('Email', 'Enter Email Used for Arena Login')
        password = tkSimpleDialog.askstring('Password', 'Enter Arena Password')
        if not email or not password:
            show_error('Value Error', 'Please Enter Email and Password')
        with open('C:\\Users\\%s\\Documents\\login.dat'%os.environ.get('USERNAME'), 'w') as f:
            f.write(email+'\n'+xor_crypt_string(password, encode=True))

    name_hash = {'adhurjaty': 'Anil Dhurjaty', 'rsleiman': 'Richard Sleiman',
                     'tkanusky': 'Thomas Kanusky', 'pneilson': 'Peter Neilson'}

    engineer = name_hash[email.split('@')[0]]
    
    return [email, password, engineer]
    

def do_something():
    print 'asdfad'
    
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

    try:
        error_info = br.find_element_by_id('loginErrorInfo')
        os.remove('C:\\Users\\%s\\Documents\\login.dat'%os.environ.get('USERNAME'))
        br.quit()
        show_error('Login Error', error_info.find_element_by_tag_name('li').text)
    except:
        pass
    
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
        show_error('Part Exists Error', edit_errors[0].find_element_by_tag_name('li').text)

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
        #pass
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

    try: #if there are no items that match part number
        table = br.find_element_by_id('PbopHeader')
        show_error('Non matching error', 'No item matches that part number')
    except:
        pass
    
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


root = Tk.Tk()
myapp = MyApp(root)
root.mainloop()



