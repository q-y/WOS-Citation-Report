# coding:utf-8

from selenium import webdriver
import time
from selenium.webdriver.support.wait import WebDriverWait
#from selenium.webdriver.support.select import Select
import Queue
import threading
import contextlib
import sys
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
import traceback
import datetime
import os


#################################
# Parameter Settings #
marked_list_url = "http://apps.webofknowledge.com/ViewMarkedList.do?action=Search&product=WOS&SID=XXXXXXXXXXXXXXX&mark_id=UDB&search_mode=MarkedList&colName=WOS" # the url for the marked article list
author_name = 'FamilyNameFirstName'.upper() # Format: 'FamilyNameFirstName'.upper(), the author's self name will be used for distinguishing self citations
timeout_s = 60 # driver.implicitly_wait value in second, typically wait the browser for 1 min
# timestr=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
# output_detailed_filename = 'output_detailed_'+timestr+'.txt'
output_detailed_filename = 'output_detailed.txt'
par_pool_size = 10 # the total number of browsers running parallelly
retry_num = 5 # the maximum error retry number
default_window_size = (1000,600,) # the window size for each browser, too small will lead to element invisible error
# End Parameter Settings #
#################################




def produce_driver():
    
    driver = webdriver.Firefox()
    
    driver.set_window_size(*default_window_size)
    return driver


def preserve_one_tab(driver):
    handles = driver.window_handles
    for i in range(len(handles)-1):
        driver.switch_to.window(handles[i])
        driver.close()
    driver.switch_to.window(handles[-1])


class ThreadPool(object):
    StopEvent = object()

    def __init__(self, max_num):
        self.q = Queue.Queue()
        self.max_num = max_num

        self.terminal = False
        self.generate_list = []
        self.free_list = []
        self.result_list = []
        self.job_id = 0


    def init_handle(self):
        for retry_i in range(retry_num):
            try:
                driver = produce_driver()
                driver.implicitly_wait(timeout_s)
            except:
                print "\n[RETRY init_handle: "+str(retry_i)+"]\n" 
                sys.stdout.write('.'*(self.q.qsize())+'\r')
                continue
            break
        else:
            s = sys.exc_info()
            print "===========Error '%s' happened on line %d.===========" % (s[1], s[2].tb_lineno)
            traceback.print_exc()
            print "===========ThreadPool: init_handle ERROR (Maximum Retry)==========="
            raise Exception("!!! ThreadPool: init_handle ERROR (Maximum Retry).")
            os._exit(1)
        return driver


    def deinit_handle(self,driver):
        # driver.close()
        driver.quit()


    def add_job(self, func, args):
        if len(self.free_list) == 0 and len(self.generate_list) < self.max_num:
            self.generate_thread()
        w = (func, args, self.job_id,)
        self.result_list.append([])
        self.job_id += 1
        self.q.put(w)


    def generate_thread(self):
        t = threading.Thread(target=self.call)
        t.setDaemon(True)
        t.start()


    def call(self):
        current_thread = threading.currentThread()
        self.generate_list.append(current_thread)
        handle = self.init_handle()

        event = self.q.get()
        while event != ThreadPool.StopEvent:
            # print current_thread, event
            func, arguments, job_id = event
            for retry_i in range(retry_num):
                try:
                    result = func(handle, *arguments)
                    self.result_list[job_id] = result
                except:
                    # self.q.put((func, arguments, job_id,))
                    print "\n[RETRY "+str(retry_i)+"]" #Error '%s' happened on line %d." % (s[1], s[2].tb_lineno)
                    traceback.print_exc()
                    preserve_one_tab(handle)
                    handle.maximize_window()
                    # len(self.generate_list) should be the number of ThreadPool.StopEvent in q
                    sys.stdout.write('.'*(self.q.qsize()-len(self.generate_list))+'\r')
                    continue
                break
            else:
                s = sys.exc_info()
                print "===========Error '%s' happened on line %d.===========" % (s[1], s[2].tb_lineno)
                traceback.print_exc()
                print "==========="+str(arguments)+"==========="
                print "===========ThreadPool: Maximum retry number arrived.==========="
                raise Exception("!!! Maximum retry number arrived.")
                os._exit(1)
            handle.set_window_size(*default_window_size)
            if self.terminal:  # False
                event = ThreadPool.StopEvent
            else:
                with self.worker_state(self.free_list, current_thread):
                    event = self.q.get()
        else:
            self.generate_list.remove(current_thread)
            self.deinit_handle(handle)


    @contextlib.contextmanager
    def worker_state(self, x, v):
        x.append(v)
        try:
            yield
        finally:
            x.remove(v)


    def close(self):
        num = len(self.generate_list)
        while num:
            self.q.put(ThreadPool.StopEvent)
            num -= 1


    def wait_all_complete(self):
        while len(self.generate_list) != 0:
            self.generate_list[0].join()


    def get_results(self):
        return self.result_list


    def terminate(self):
        self.terminal = True
        self.q.queue.clear()
        num = len(self.generate_list)
        while num:
            self.q.put(ThreadPool.StopEvent)
            num -= 1
        # for item in self.generate_list:
        #     item.stop()
        self.wait_all_complete()




def long_sleep():
    time.sleep(2)


def short_sleep():
    time.sleep(0.5)


def wait_for_new_page(driver):
    long_sleep()
    WebDriverWait(driver, timeout_s).until(
        lambda driver: driver.execute_script("return document.readyState;") == "complete")
    time.sleep(0.1)


# find the substring index at the "findCnt" occurances of "subStr"
def findNStr(string, subStr, findCnt):
    listStr = string.split(subStr,findCnt)  #2nd para is maxsplit: to split the string into maximum of provided number of times
    # print listStr
    if len(listStr) <= findCnt:    
        return -1
    return len(string)-len(listStr[-1])-len(subStr)
    #len(listStr[-1]) is the length for the last collection, has to also minus len(subStr) itself


def isElementExist(driver, parent, element):
    try:
        driver.implicitly_wait(0)
        parent.find_element_by_xpath(element)
        driver.implicitly_wait(timeout_s)
        return True
    except:
        driver.implicitly_wait(timeout_s)
        return False


def get_all_records(driver):
    all_record_num = driver.find_element_by_id("hitCount.top").text
    driver.find_element_by_xpath("//button[contains(.,'Print')]").click()
    # wait_for_new_page(driver)
    short_sleep()
    driver.find_element_by_xpath("(//input[@id='numberOfRecordsRange'])").click()
    driver.find_element_by_xpath("(//input[@id='markFrom'])").clear()
    driver.find_element_by_xpath("(//input[@id='markFrom'])").send_keys("1")
    driver.find_element_by_xpath("(//input[@id='markTo'])").clear()
    driver.find_element_by_xpath("(//input[@id='markTo'])").send_keys(all_record_num)
    driver.find_element_by_xpath("//select[@id='bib_fields']/following-sibling::*[1]").click()
    #Select(driver.find_element_by_xpath("//select[@id='bib_fields']")).select_by_index(0)
    driver.find_elements_by_xpath("//ul")[-1].find_element_by_tag_name("li").click()
    driver.find_element_by_xpath("//*[@title='Print']").click()
    handles = driver.window_handles  # obatin all the handles in a list
    driver.switch_to.window(handles[1])
    wait_for_new_page(driver)
    record_list = []
    for i in range(int(all_record_num)):
        curr_i = i % 50 + 1
        string = driver.find_element_by_xpath("//*[@id='printForm']/table[" + str(curr_i + 1) + "]").text
        string = string[0:findNStr(string, '\n', 5)]
        record_list.append(string)
        if curr_i==50 and i!=int(all_record_num)-1:
            driver.find_element_by_xpath("//*[@title='Next Page']").click()
            WebDriverWait(driver, timeout_s).until(
                lambda driver: driver.find_element_by_xpath("//b[text()='Record "+str(i+2)+" of "+all_record_num+"']"))
            wait_for_new_page(driver)
    driver.close()
    driver.switch_to.window(handles[0])  # switch to the beginning window
    return record_list



def rearrange_list_count(list):
    list_num = len(list)
    for curr_i in range(list_num):
        strs = list[curr_i].splitlines(True)
        strs[0] = 'Record ' + str(curr_i + 1) + ' of ' + str(list_num) + '\n'
        list[curr_i] = ''.join(strs)



def fetch_a_record(driver, id, url):
    driver.get(url)
    driver.find_element_by_xpath("//span[@id='citationScoreCard']/div[2]/p[9]/a[2]").click()
    a_WOS_cite=[]
    a_WOS_cite.append(driver.find_element_by_xpath("//*[@id='CAScorecard_count_WOS']").text) #WoS
    if isElementExist(driver, driver, "//*[@id='CAScorecard_count_WOSCLASSIC']/a"):
        a_SCI_cite=[]
        item_tmp=driver.find_element_by_xpath("//*[@id='CAScorecard_count_WOSCLASSIC']/a")
        #SCI_total_num=item_tmp.text     resolve neq
        item_tmp.click()
        wait_for_new_page(driver)
        all_SCI_list=get_all_records(driver)
        a_SCI_cite.append(str(len(all_SCI_list)))  # SCI total num
        by_others_list=[]
        by_self_list = []
        for item in all_SCI_list:
            strs=item.splitlines(True)
            au=''.join([x for x in strs[2] if x.isalpha()]).upper()
            if au.find(author_name)==-1:
                by_others_list.append(item)
            else:
                by_self_list.append(item)
        rearrange_list_count(by_others_list)
        rearrange_list_count(by_self_list)
        a_SCI_cite.append(by_others_list)
        a_SCI_cite.append(by_self_list)
        a_WOS_cite.append(a_SCI_cite)  # SCI
    else:
        a_WOS_cite.append([]) # SCI
    sys.stdout.write('*')
    return id, a_WOS_cite


def get_cite_records():
    sys.stdout.write('Getting record list ')
    record_list = []
    driver = produce_driver()
    driver.implicitly_wait(timeout_s)
    driver.get(marked_list_url)
    str_tmp=driver.find_element_by_xpath("//*[@id='output_form']/div[2]/span/span[1]").text
    all_record_num = int(str_tmp[0:str_tmp.index(' ')])
    sys.stdout.write('['+str(all_record_num)+'] ...   ')
    if driver.find_element_by_xpath("//*[@id='select2-selectPageSize_bottom-container']").text.find("50")==-1:
        driver.find_element_by_xpath("//*[@id='select2-selectPageSize_bottom-container']").click()
        driver.find_element_by_xpath("//*[@id='select2-selectPageSize_bottom-results']/li[3]").click()
        wait_for_new_page(driver)
    #driver.find_element_by_xpath("//input[@name='formatForPrint']").click()
    driver.find_element_by_xpath("//button[contains(.,'Print')]").click()
    handles = driver.window_handles  # obatin all the handles in a list
    driver.switch_to.window(handles[1])
    wait_for_new_page(driver)
    for i in range(all_record_num):
        curr_i = i % 50 + 1
        string = driver.find_element_by_xpath("//*[@id='printForm']/table[" + str(curr_i + 1) + "]").text
        record_list.append(string)
        if curr_i == 50 and i!=all_record_num-1:
            driver.find_element_by_xpath("//*[@title='Next Page']").click()
            WebDriverWait(driver, timeout_s).until(
                lambda driver: driver.find_element_by_xpath(
                    "//b[text()='Record " + str(i + 2) + " of " + str(all_record_num) + "']"))
            wait_for_new_page(driver)
    driver.close()
    driver.switch_to.window(handles[0])
    sys.stdout.write('Done.\n')

    sys.stdout.write('Getting citation results ('+str(all_record_num)+' items)...\n')
    sys.stdout.write('.'*all_record_num+'\r')
    # [[0 for col in range(5)] for row in range(3)]
    cite_result = [[0 for col in range(2)] for row in range(all_record_num)]  # each entry: [a_WOS_cite,isHighlyCited[0/1]]
    pool = ThreadPool(par_pool_size)
    for i in range(all_record_num):
        curr_i=i%50+1
        if isElementExist(driver,driver,"//*[@id='RECORD_"+str(i+1)+"']/div[5]/div[1]/a"):
            url=driver.find_element_by_xpath("//*[@id='RECORD_" + str(i+1) + "']/div[5]/div[1]/a").get_attribute("href")
            pool.add_job(func=fetch_a_record, args=(i,url,))
            # is highly cited
            #tmp_record = driver.find_element_by_xpath("//*[@id='RECORD_"+str(i+1)+"']/div[5]")
            #if isElementExist(driver, tmp_record, "//div[contains(@id,'div_highlyCitedBadge')]"):
            if isElementExist(driver, driver, "//*[@id='div_highlyCitedBadge_" + str(i+1) + "']"):
                cite_result[i][1] = True
            else:
                cite_result[i][1] = False
        else:
            cite_result[i][1] = False
            cite_result[i][0] = []
            sys.stdout.write('*')
        if curr_i==50 and i!=all_record_num-1:
            driver.find_element_by_xpath("//*[@title='Next Page']").click()
            wait_for_new_page(driver)
    pool.close()
    pool.wait_all_complete()
    for id, a_WOS_cite in pool.get_results():
        cite_result[id][0] = a_WOS_cite
    driver.quit()
    sys.stdout.write('\nDone.\n')
    return record_list, cite_result


def write_file(record_list, cite_result):
    file_detailed = open(output_detailed_filename, 'w')
    str_detailed = ''
    for i in range(len(record_list)):
        str_detailed = str_detailed + record_list[i]+'\n'
        if len(cite_result[i][0])==0:
            str_detailed = str_detailed + 'WOS Citation: 0'
        else:
            a_WOS_cite=cite_result[i][0]
            str_detailed = str_detailed + 'WOS Citation: '+a_WOS_cite[0]+'   '
            if len(a_WOS_cite[1])==0:
                str_detailed = str_detailed + 'SCI Citation: 0'
            else:
                a_SCI_cite=a_WOS_cite[1]
                str_detailed = str_detailed + 'SCI Citation: ' + a_SCI_cite[0] + ' (By others: '+str(len(a_SCI_cite[1]))+', self: '+str(len(a_SCI_cite[2]))+')'
                assert(int(a_SCI_cite[0])==len(a_SCI_cite[1])+len(a_SCI_cite[2]))
                if cite_result[i][1]:
                    str_detailed = str_detailed + '\n*** Highly Cited Paper ***'
                if len(a_SCI_cite[1])!=0:
                    str_detailed = str_detailed + '\nBy others:\n\t'
                    lines = ''
                    for line in a_SCI_cite[1]:
                        lines=lines + line + '\n'
                    lines=lines[:-1].replace('\n','\n\t')
                    str_detailed = str_detailed + lines
                if len(a_SCI_cite[2])!=0:
                    str_detailed = str_detailed + '\nBy self:\n\t'
                    lines = ''
                    for line in a_SCI_cite[2]:
                        lines=lines + line + '\n'
                    lines=lines[:-1].replace('\n','\n\t')
                    str_detailed = str_detailed + lines
        str_detailed = str_detailed + '\n\n'
    file_detailed.write(str_detailed)
    file_detailed.close()


def main():
    record_list, cite_result = get_cite_records()
    sys.stdout.write('Writing files...  ')
    write_file(record_list, cite_result)
    sys.stdout.write('Done.\n')
    raw_input()


if __name__ == "__main__":
    main()
