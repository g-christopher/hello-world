def strip_address(test_str):
    dummy = ''
    dummy2= ''
    skip1c = 0
    skip2c = 0
    for i in test_str:
        
        if i == '(':
            skip1c += 1
        elif i == ')' and skip1c > 0:
            skip1c -= 1
        elif skip1c == 0:
            dummy += i
        else:
            dummy2 +=i
            ###print 'authors are:', dummy, 'address is:', dummy2
    return dummy2
        
        ### function for striping {}
def strip_email(test_str):
    dummy = ''
    dummy2= ''
    skip1c = 0
    for i in test_str:
        
        if i == '{':
            skip1c += 1
        elif i == '}' and skip1c > 0:
            skip1c -= 1
            break
        elif skip1c == 0:
            dummy += i
        else:
            dummy2 +=i
            #print 'authors are:', dummy, 'email is:', dummy2
    
    return dummy2

def strip_language(test_str):
    dummy = ''
    dummy2= ''
    skip1c = 0
    skip2c = 0
    for i in test_str:
        
        if i == '(':
            skip1c += 1
        elif i == ')' and skip1c > 0:
            skip1c -= 1
        elif skip1c == 0:
            dummy += i
        else:
            dummy2 +=i
            ###print 'title is:', dummy, 'language is:', dummy2
    return dummy2

def strip_title(test_str):
    dummy = ''
    dummy2= ''
    skip1c = 0
    skip2c = 0
    for i in test_str:
        
        if i == u'\u201C' or i == u'\u2018':
            skip1c += 1
        elif (i == u'\u201D' or i == u'\u2019') and skip1c > 0:
            skip1c -= 1
        elif skip1c == 0:
            dummy += i
        else:
            dummy2 +=i
            ###print 'title is:', dummy, 'language is:', dummy2
    return dummy2        

def strip_original(test_str):
    dummy = ''
    dummy2= ''
    skip1c = 0
    skip2c = 0
    test_str=test_str.replace('<<','{',1)
    test_str=test_str.replace('>>','}',1)
    print test_str
    for i in test_str:
        
        if i == '{':
            skip1c += 1
        elif i == '}' and skip1c > 0:
            skip1c -= 1
        elif skip1c == 0:
            dummy += i
        else:
            dummy2 +=i
            #print 'orig is:', dummy, 'other is:', dummy2
    #print 'Title:', dummy, 'Original:', dummy2        
    return dummy2

def aucleanup(test_str):
    dummy=test_str
    autitlelist=['Prof.', 'Professor', 'Dr.', '[']
    for autitle in autitlelist:
        if autitle in test_str:
            dummy=test_str.split(autitle).pop().strip()
            
    return dummy

def xmlstrip(test_str,name):

    spl_name='<'+name
    dummy=test_str.split(spl_name).pop()
    dummy=dummy.split('<')[0]
    ##print '0000', dummy
    dummy=dummy.split('>').pop()
    ##print '1111', dummy
    dummy=dummy.strip()

    return(dummy)
    

### function for fixing dates. not working ###


def datemaker(datein):
    
    
    Months=['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    Months_abbr=['Jan', 'Feb', 'Mar', 'Apr', 'Jun', 'Jul', 'Aug', 'Sep', 'Sept', 'Oct', 'Nov', 'Dec']
    nomonth=1
    months_list=Months+Months_abbr
    for month in months_list:
        if month in datein:
            nomonth=1
            
    datin_list=datein.split(' ')
    if len(datein)==3:
        dateout=datein
    else:
        pass
            
    return dateout
                                                                                                                                                                                
