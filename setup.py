from distutils.core import setup
import py2exe
import shutil

setup(console=['arena_automate.py'],
      options={
            "py2exe":{
                    "skip_archive": True,
                    "unbuffered": True,
                    "optimize": 2
            }
    })

shutil.copytree('chrome', 'dist\\chrome')
