## Company: Yang bigdagta Tech Ltd, CN
## Occupation: Junior Assi. Engineer  助理工程师
## Goal: 
Develope a python script to extract TEXT elements (ex. text boxes, speaker notes, charts, columns) from variety of ppt/excel/.docx into database, Thus setup other NLP engineers (算法工程师) to train A GPT4 model to give feedback based on the data within the database
## Case Analysis:
The business department has many knowledge based documents(lesson-learn, guide, specification, tutorials, etc.) They are all in word and ppt format. 
However, when they want to search the information, but it is difficult since file data is qualitive and abstract, the file name is non-standard and they are require to open file one by one to search the details matching their intent. Sometimes, they couldn't even find it as the text is very spelling sensitive.

## Responsibilities
 1. To convert the all ppt and word file to a txt format.
 2. Import documents into at able of mysql database, where they can search though the table conviently.
 3. Allow future engineers to use these text to train AI model like GPT, where the users could easily to get the question-answer.

### MYSQL Guidethrough
In mysql environment 
create  dabase  office_text
use office_text 

3
CREATE TABLE ppt_text(
  file_number int(11) NOT NULL AUTO_INCREMENT,
  file_name  varchar(512) NOT NULL,
  page_num int(5) NOT NULL,
  page_content varchar(5000) DEFAULT NULL,
  load_time datetime NOT NULL,
  PRIMARY KEY (`file_number`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

Note: The first field of the first column is auto-incremented.

![Final Product](https://github.com/frankqgu/textextract/assets/130730924/8c62c57f-6e14-4950-9f9d-1ccf06ee1067)
