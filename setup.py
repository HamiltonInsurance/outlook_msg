import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="outlook_msg",
    version="1.0.0",
    author="Hamilton Insurance Group",
    author_email="opensource@hamiltoninsurancegroup.com",
    description="Read contents, metadata and attachments from Outlook Message files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/HamiltonInsurance/outlook_msg",
    install_requires=["compoundfiles~=0.3", "compressed_rtf~=1.0.6"],
    python_requires=">=3.6, <4",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: Financial and Insurance Industry",
        "Topic :: Communications :: Email",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
)