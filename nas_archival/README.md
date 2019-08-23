# NAS Archival Script
To generate NAS Archival .pdf and .docx from Webpage

## Requirements
1. Python >= 3.3 https://www.python.org/downloads/
2. LibreOffice >= 6.0 https://www.libreoffice.org/download/download/
3. Write, Remove, Create Files and Directories Permissions
4. Windows requires Git Bash https://git-scm.com/downloads
5. No proxy if have set before running

## How to Use
### 1. Generate for year and category

#### Run Script
In Windows or Linux:
```
python3 nas.py --year <year> --category <category> --follow-related-links <True/False> (optional) 
```

In MacOS
```
OBJC_DISABLE_INITIALIZE_FORK_SAFETY=YES python3 nas.py --year <year> --category <category> --follow-related-links <True/False> (optional) 
```

*e.g. python3 nas.py --year 2019 --category news-releases*
#### Script Parameters
**year** - Year of pages

**category** - Category of pages

*VALID CATEGORIES - news-releases, speeches, others*

*(optional default is True)* **follow-related-links** - Generate for related resources?

#### Output
**Directory:** *\<category\>/\<year>\/\<month string\>/\<yyyymmdd\>*

*e.g. news-releases/2019/march/20190330, news-releases/2019/march/20190330_1*

**Files:**
  - pdf, MINDEF_20190330001.pdf
  - docx，MINDEF_20190330001.docx
  - (if any) images, MINDEF_20190330001_IMG_0.png, MINDEF_20190330001_IMG_1.png,....
  - (if any) related pdfs, MINDEF_20190330002.pdf, MINDEF_20190330003.pdf
  - (if any) related docx, MINDEF_20190330002.docx, MINDEF_20190330003.docx
  - (if any) related images, MINDEF_20190330003_IMG_0.PNG,... 
  - debug.txt: To retrieve info for csv and for debugging
  - details.txt: Title and Link of Articles fetched
  - error.txt: Error log if its unsuccessful, have to do manual editing

### 2. Generate for a list of urls

#### Run Script
In Windows or Linux:
```
python3 nas.py --urls <list of urls seperated by comma> --follow-related-links <True/False> (optional) 
```

In MacOS
```
OBJC_DISABLE_INITIALIZE_FORK_SAFETY=YES python3 nas.py --urls <list of urls seperated by comma> --follow-related-links <True/False> (optional) 
```

*e.g. python3 nas.py --urls url1, url2, url3*

##### Running a list of urls from file seperated by line breaks
Windows requires Linux tools can be run via Git Bash
```
python3 nas.py --urls `cat <filename>|tr '\n' ','` --follow-related-links <True/False> (optional) 
```

#### Script Parameters
**urls** - list of urls seperated by comma

*(optional default is True)* **follow-related-links** - Generate for related resources?

#### Output
**Directory:** *manual/\<year\>/\<month string\>/\<yyyymmdd\>*

*e.g. manual/2019/march/20190330, manual/2019/march/20190330_1*

**Files:**
  - pdf, MINDEF_20190330001.pdf
  - docx，MINDEF_20190330001.docx
  - (if any) images, MINDEF_20190330001_IMG_0.png, MINDEF_20190330001_IMG_1.png,....
  - (if any) related pdfs, MINDEF_20190330002.pdf, MINDEF_20190330003.pdf
  - (if any) related docx, MINDEF_20190330002.docx, MINDEF_20190330003.docx
  - (if any) related images, MINDEF_20190330003_IMG_0.PNG,... 
  - debug.txt: To retrieve info for csv and for debugging
  - details.txt: Title and Link of Articles fetched
  - error.txt: Error log if its unsuccessful, have to do manual editing

### 3. Generate for url
#### Run Script
In Windows or Linux:
```
python3 nas.py --url <url> --follow-related-links <True/False> (optional) 
```

In MacOS
```
OBJC_DISABLE_INITIALIZE_FORK_SAFETY=YES python3 nas.py --url <url> --follow-related-links <True/False> (optional) 
```

*e.g. python3 nas.py --url url1*

#### Script Parameters
**url** - url of article

*(optional default is True)* **follow-related-links** - Generate for related resources?

#### Output
**Directory:** *manual/\<year\>/\<month string\>/\<yyyymmdd\>*

*e.g. manual/2019/march/20190330, manual/2019/march/20190330_1*

**Files:**
  - pdf, MINDEF_20190330001.pdf
  - docx，MINDEF_20190330001.docx
  - (if any) images, MINDEF_20190330001_IMG_0.png, MINDEF_20190330001_IMG_1.png,....
  - (if any) related pdfs, MINDEF_20190330002.pdf, MINDEF_20190330003.pdf
  - (if any) related docx, MINDEF_20190330002.docx, MINDEF_20190330003.docx
  - (if any) related images, MINDEF_20190330003_IMG_0.PNG,... 
  - debug.txt: To retrieve info for csv and for debugging
  - details.txt: Title and Link of Articles fetched
  - error.txt: Error log if its unsuccessful, have to do manual editing

## Notes
Related More Resources Supported are
- News Releases
- Speeches
- Others
Rest will be ignored such as reply to mq, cyberpioneer/army/navy/airforce articles

## Possible Problems
- Unhandled Tags: font, style, figure, aside
- Unhandled Attributes: float, margin, padding, width, height
- Image expanding more than resolution
- Portrait Images takes up a page
- Table widths does not follow html
- Some Images are not able to be added due to the source EXIF data, Have to be added manually
- Network Issues
- Unsolved Table/Other errors (report to me i will fix)
- nbsp; removed thus combining words into 1 word
- Duplicate files conversion due to different links but same file, Have to manually cleanup and change More part in docx and pdf. 
- HDD no space



