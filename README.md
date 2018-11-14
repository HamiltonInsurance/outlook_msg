# outlook_msg

[![Build Status](https://travis-ci.org/HamiltonInsurance/outlook_msg.svg?branch=master)](https://travis-ci.org/HamiltonInsurance/outlook_msg)

`outlook_msg` is a Python library by [Hamilton Group](http://www.hamiltongroup.com/) to process the .msg files that 
Users can export from Outlook. It is very common for users in organizations that use Outlook to archive data in this 
format. For example, at Hamilton we see these files store in relation to deals we write. If we want to do automatic
processing in a way that feels most natural to end users we need to extract data from these files.

This library is built on top of the excellent [compoundfiles](https://pypi.org/project/compoundfiles/) library, without
which none of this would be possible.
 

## Getting Started

Install using pip:
 
`pip install outlook_msg`
 
## Usage

To open an email:

```python
from outlook_msg import Message

with open('file.msg') as msg_file:
    msg = Message(msg_file)
    
# Contents are the plaintext body of the email
contents = msg.body

# Attachments can be read and saved like so
first_attachment = msg.attachments[0]
with first_attachment.open() as attachment_fp, open(first_attachment.filename, 'wb') as output_fp:
    output_fp.write(attachment_fp.read())
``` 

## Running the tests

We use [pytest](https://docs.pytest.org/en/latest/) to run our tests but you are best to use tox so you can test on all
supported Python versions and will ensure a clean environment.  