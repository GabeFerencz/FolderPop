# Gabe Ferencz
# FolderPop - Moves all contents of a folder into its parent directory and 
#             delete the empty folder.
# To install, run the script, which will add: 
# "C:\Python26\python.exe" "_path_to_script_\FolderPop.py" "%1"
# to the registry in:
# HKEY_CLASSES_ROOT\Folder\Shell\Folder Pop\command

def FolderPop(folder):
    '''Move folder contents to parent directory and delete empty folder.'''
    import os,glob,shutil
    # First rename the directory to avoid duplicate directory names
    os.renames('./'+folder,'./TEMP_FOLDER_POP_FOLDER')
    # Now, move all of the contents of the folder to the parent directory
    for item in glob.glob('./TEMP_FOLDER_POP_FOLDER/*'):
        shutil.move(item,'.')
    # Finally, delete the temporary folder
    os.rmdir('./TEMP_FOLDER_POP_FOLDER')
    return

def Install(menu_name='Folder Pop'):
    '''Install registry entry for adding Folder Pop to context menu.'''
    import _winreg, os
    
    q = ("WARNING: You should back up your registry before making changes!" + 
         "\n\nAre you sure you want me to edit your registry now?")
    if YesNoDialog(q):
        keyVal = 'Folder\\Shell\\' + menu_name + '\\command'
        try:
            key = _winreg.OpenKey(_winreg.HKEY_CLASSES_ROOT, 
                                    keyVal, 
                                    0, 
                                    _winreg.KEY_ALL_ACCESS)
        except WindowsError:
            key = _winreg.CreateKey(_winreg.HKEY_CLASSES_ROOT, keyVal)
        regEntry = (r'"C:\Python26\python.exe" "' + os.getcwd() + 
                    r'\FolderPop.py" "%1"')
        _winreg.SetValueEx(key, '', 0, _winreg.REG_SZ, regEntry)
        _winreg.CloseKey(key)
    return

def YesNoDialog(question, title = 'Folder Pop Confirmation Box'):
    '''Simple yes/no dialog box that uses win32 extensions if available.'''
    try:
        import win32ui as wu
        import win32con as wc
        ans = wu.MessageBox(question, title, 
                            (wc.MB_YESNO + wc.MB_ICONQUESTION)) == wc.IDYES
    except ImportError:
        ans = raw_input(question).lower() in ['y','yes']
    return ans

if __name__ == "__main__":
    import sys, os
    
    if len(sys.argv) == 1:
        Install()
    else:
        # The right-clicked folder's full path is the second input argument
        full_path = sys.argv[1]
        # Extract the folder name from the full path, the current directory
        #    is the parent directory of that folder
        folder_name = full_path.split('\\')[-1] 
    
    if len(sys.argv) == 2:
        q = ('Are you sure you want to pop:\n\n' + full_path + 
            '    to:\n' + os.getcwd())
        if YesNoDialog(q):
            FolderPop(folder_name)
    elif len(sys.argv) == 3:
        # Add a dummy input to registry entry to quiet the confirmation
        FolderPop(folder_name)
