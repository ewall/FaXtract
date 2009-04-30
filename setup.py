from distutils.core import setup
import py2exe

setup(
    name         = "FaXtract",
    version      = "0.3",
    author       = "Eric Wallace",
    author_email = "eric.wallace@atlanticfundadmin.com",
    description  = "Extracts PDF faxes out of IMAP mailboxes",
    
    console      = ['FaXtract.py']

    )

"""
        options      = {
                            "py2exe":{
                                "unbuffered": True,
                                "optimize": 2,
                                "includes": ["email"]
                            }
                       }
"""
