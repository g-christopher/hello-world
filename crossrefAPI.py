#!/usr/bin/env python
from __future__ import division
import requests
import xmltodict
from strippers import xmlstrip
from crossref.restful import Works

def Getdoiplus(markupdict):
    #### Get DOI of journal article from CrossRef XML queries
    #### Uses crossref.restful Works python module
    #### Inputs are Title and First Author
    #### Output is DOI

    print 'Attempting to get DOI from CrossRef'
    works=Works()
    titlestr=markupdict['Title']
    ###Find First Author Surname
    if ' ' in markupdict['FirstAuthor']:
        dummydict=markupdict['FirstAuthor'].split(' ')
        if ',' in markupdict['FirstAuthor']:
            authorstr=dummydict[0].strip()
            authorstr=authorstr.replace(',','',1)
        else:
            authorstr=dummydict.pop().strip()
    else:
        authorstr=markupdict['FirstAuthor']
    if '-' in authorstr:
        authorstr=authorstr.split('-').pop()
        
    #print markupdict['FirstAuthor'], 'authorst=',authorstr

    ### Query does not do exact phrase matching - find word in title that returns fewest results
    titlestr=''
    leastquerynumber=999999999
    #print markupdict['Title'].split(' ')
    #dummy=markupdict['Title'].replace('\u\xa0','')
    for word in markupdict['Title'].split(' '):
        if len(word)>5:
            #print word, works.query(title=word, author=authorstr).count()
            if works.query(title=word, author=authorstr).count()<leastquerynumber and works.query(title=word, author=authorstr).count()>0:
                titlestr=word
                #print 'titlestr is', word
            
    if titlestr=='': titlestr=markupdict['Title']
    #print markupdict['Title'], 'titlestr=', titlestr
    
    DOIstr='10.1016/s0022-3115(98)00906-4'
    
    print 'number of titles is:',  works.query(title='Uranium dioxide', author='Bae' ).count()
    #print 'number of titles is:',  works.query(title=titlestr, author='Bae' ).count()

    #print 'doi no. titles:', works.query(DOI=DOIstr)
    doi=''
    ### Perform query and find exact or partial title matches
    print 'Querying Title=%s Author=%s number of titles is: %d' % (titlestr, authorstr, works.query(title=titlestr, author=authorstr).count())

    for item in works.query(title=titlestr, author=authorstr):
        #for item in works.query(title='Uranium dioxide', author='Bae' ):
        #for item in works.query(DOI=DOIstr):    
        #print item['title'][0]
        if markupdict['Title'].lower()==item['title'][0].lower():
            print '**** exact match ****'
            print '[1]', item['title'][0]
            print '[2]', markupdict['Title']
    
            doi=item['DOI']
        elif abs(len(item['title'][0].split(' '))-len(markupdict['Title'].split(' ')))==0: ### Word by word matching
            #dummytitle1=item['title'][0].replace('(',"",1).replace(')',"",1).replace('/',"",1).replace('.',"",1).replace(',',"",1)
            n_match=0
            for i in range (len(item['title'][0].split(' '))):
                if item['title'][0].split(' ')[i].strip().lower() == markupdict['Title'].split(' ')[i].strip().lower():
                    n_match=n_match+1
            #if abs(len(item['title'][0].split(' '))-n_match)<3:'
            if n_match/(len(item['title'][0].split(' '))*1.0)>0.7 and abs(len(item['title'][0].split(' '))-n_match)<3:
                doi=item['DOI']
                print '**** partial match ****', n_match, ' of', len(item['title'][0].split(' ')), 'matches. Proportion:', n_match/(len(item['title'][0].split(' '))*1.0)
                print '[1]', item['title'][0]
                print '[2]', markupdict['Title']


            #print 'matches=', n_match, 'total=', len(item['title'][0].split(' '))

        ##    print item['DOI']
    
    ##for item in works.sample(2):
    ##    print (item['title'])
    ##    print item['DOI']
    if doi=='':
        print '**** No Match found ****\n'
    else:
        pass
        #print doi
        
    print '*************\n'
    #print item.keys()



    return(doi)

def GetFullAuthors(doi):
    print 'Getting full authors'
    myemail='grant.christopher@kcl.ac.uk'
    url='https://doi.crossref.org/servlet/query?pid=%s&format=unixref&id=%s' % (myemail, doi)
    r=requests.get(url)
    xmlstr= r.text
    #print xmlstr
    #authorslist=xmlstrip(xmlstr, 'contributors')
    #authorslist=xmlstrip.split('<Contributors').pop()
    xmlstr=xmlstr.strip('\r\n')
    print xmlstr
    authorsdict=''
    return authorsdict


'''
url2='https://doi.crossref.org/servlet/query?pid=grant.christopher@kcl.ac.uk&id=10.1577/H02-043'
##rr=requests.get(url2)
##print rr.text





###url='http://doi.crossref.org/servlet/query?pid=grant.christopher@kcl.ac.uk&format=unixref&qdata=<?xml version=\"1.0\"?><query_batch version=\"2.0\" xsi:schemaLocation=\"http://www.crossref.org/qschema/2.0 http://www.crossref.org/qschema/crossref_query_input2.0.xsd\" xmlns=\"http://www.crossref.org/qschema/2.0\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"> <head> <email_address>grant.christopher@kcl.ac.uk</email_address> <doi_batch_id>ABC_123_fff</doi_batch_id> </head> <body> <query enable-multiple-hits=\"false\" secondary-query=\"author-title\" key=\"key1\"><article_title match=\"fuzzy\"> Beneficiation of Saudi phosphate ores by column flotation technology</article_title> <author search-all-authors=\"true\"> Al-Fariss</author> </query> </body> </query_batch>'

Title='Beneficiation of Saudi phosphate ores by column flotation technology'
FirstAuthor='Al Fariss'

url='http://doi.crossref.org/servlet/query?pid=grant.christopher@kcl.ac.uk&format=XSD_XML&qdata=<?xml version=\"1.0\"?><query_batch version=\"2.0\" xsi:schemaLocation=\"http://www.crossref.org/qschema/2.0 http://www.crossref.org/qschema/crossref_query_input2.0.xsd\" xmlns=\"http://www.crossref.org/qschema/2.0\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"> <head> <email_address>grant.christopher@kcl.ac.uk</email_address> <doi_batch_id>ABC_123_fff</doi_batch_id> </head> <body> <query enable-multiple-hits=\"false\" secondary-query=\"author-title\" key=\"key1\"><article_title match=\"fuzzy\"> %s</article_title> <author search-all-authors=\"true\"> %s</author> </query> </body> </query_batch>' % (Title, FirstAuthor)

r=requests.get(url)
##print r.text
##teststr='<doi type'+r.text.split('<doi type').pop()
##print teststr
##teststr=teststr.split('</query')[0].strip()
##print teststr
##mydict=xmltodict.parse(r.text)
##print mydict

##str2='<crossref_result xmlns="http://www.crossref.org/qrschema/2.0" version="2.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.crossref.org/qrschema/2.0 http://www.crossref.org/schema/queryResultSchema/crossref_query_output2.0.xsd"><query_result><head><email_address>grant.christopher@kcl.ac.uk</email_address><doi_batch_id>ABC_123_fff</doi_batch_id></head><body><query key="key1" status="resolved" fl_count="0" query_mode="author-title"><!--score=79.8--><doi type="journal_article">10.1016/j.jksues.2012.05.002</doi><issn type="print">10183639</issn><journal_title>Journal of King Saud University - Engineering Sciences</journal_title><author>Al-Fariss</author><volume>25</volume><issue>2</issue><first_page>113</first_page><year media_type="print">2013</year><publication_type>full_text</publication_type></query></body></query_result></crossref_result>'

str2='<doi type'+r.text.split('<doi type').pop()
str2=str2.split('</query')[0].strip()
print xmlstrip(str2,'doi')
pubdict={}
publist=['doi', 'issn', 'journal_title', 'author', 'volume', 'issue', 'first_page', 'year', 'publication_type']
for item in publist:
    pubdict[item]=xmlstrip(str2, item)
print pubdict
'''

                                   
