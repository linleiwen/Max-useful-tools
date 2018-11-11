'''__author__:Leiwen Lin''' 
'''__maintainer__:'Leiwen Lin''' 
'''__createtime__:2018-08-08 14:23:44.421421''' 
succeed_flag = False
fail_count = 0 
while(not succeed_flag):
    try:
        #content
        succeed_flag = True
    except:
        #return initial stage content e.g:selenium go to the home page
        fail_count= fail_count +1
        print(f"failed {fail_count} times")



