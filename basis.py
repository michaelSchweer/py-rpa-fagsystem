import os
import sys
from time import sleep
from time import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

basisUrl = '<BASE URL>'

options = Options()
# options.add_argument('--headless') # kommentér denne linje ud for at få vist browseren under eksekvering
options.add_argument('--hide-scrollbars')
options.add_argument('--disable-gpu')
options.add_argument('--log-level=3')
options.add_argument('--silent')
options.add_argument('--start-maximized')
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(10)
os.system('cls')

fagsystemer = {'<SYS NAME>': '<SYS PAGE>'}

def openFagsystem(fagsystem):

    driver.execute_script('''window.open("''' + fagsystemer[fagsystem] + '''","_blank");''')
    driver.switch_to.window(driver.window_handles[-1])

    print('15 sekunders arbejde i fagsystemet...')
    sleep(15)

    driver.close()

    print('Og tilbage...')

    sys.exit()


gemteRITMer = {}
firstLoop = True
while True:
    startTid = time()
    try:
        driver.get(basisUrl)
        
        sleep(1)

        Select(driver.find_element_by_id('select_filter_type')).select_by_visible_text('Opret')
        Select(driver.find_element_by_id('select_filter_tildelt')).select_by_visible_text('Ingen')
        sleep(3)

        tbodies = driver.find_elements_by_tag_name('tbody')
        resultsTable = tbodies[6]
        resultRows = resultsTable.find_elements_by_tag_name('tr')

        # if firstLoop:
        #     print('Itererer over de første 100 sager. Dette tager ca. 30 sekunder...')
        
        for r in resultRows[:5]:
            rowID = r.get_attribute('id')
            cells = r.find_elements_by_tag_name('td')
            
            c_snRITM = int(cells[0].text)
            
            if c_snRITM in gemteRITMer:
                break
            
            c_REQ = cells[1].text
            c_Forvalt = cells[2].text
            c_Type = cells[3].text
            c_CPR = cells[4].text
            c_FuldeNavn = cells[5].text
            c_KKorg = cells[6].text
            c_Indmeldt = cells[7].text
            c_IndmeldtDato = c_Indmeldt.split(' ')[0]
            c_IndmeldtTidspunkt = c_Indmeldt.split(' ')[1]
            c_SLA = cells[8].text
            c_Ikraft = cells[9].text
            c_UdfortAf = cells[10].text
            c_UdfortDato = cells[11].text
            c_LUT = cells[12].text
            c_Att = cells[13].text
            c_Checkbox = cells[14].text

            strResult = '{} <{}> {}, KKOrg: {}'.format(c_IndmeldtTidspunkt, c_Type, c_FuldeNavn, c_KKorg)

            if c_Type == 'Opret' and c_snRITM not in gemteRITMer:
                gemteRITMer[c_snRITM] = rowID, strResult

        if firstLoop:
            os.system('cls')
        firstLoop = False
        seneste = {k: gemteRITMer[k] for k in sorted(gemteRITMer.keys(), reverse=True)[:5]}
        for k, v in seneste.items():
            print('[{}] {}'.format(k, v[1]))
            
            try:
                ActionChains(driver).move_to_element(driver.find_element_by_id(v[0])).double_click().perform()
                sleep(3)
            except Exception as e:
                print('Sagen forsvundet fra forsiden. Fortsætter...')
                continue

            bodies = driver.find_elements_by_tag_name('tbody')
            ourBody = bodies[7]
            ourRows = ourBody.find_elements_by_tag_name('tr')
            for r in ourRows:
                rowRetID = r.get_attribute('id')

                ourCells = r.find_elements_by_tag_name('td')

                d_snRITM = ourCells[0].text
                d_REQ = ourCells[1].text
                d_Forvalt = ourCells[2].text
                d_Todo = ourCells[3].text
                d_FuldeNavn = ourCells[4].text
                d_System = ourCells[5].text
                d_Rettighed = ourCells[6].text
                d_Status = ourCells[7].text
                d_Ikraft = ourCells[8].text
                d_UdfortAf = ourCells[9].text
                d_UdfortDato = ourCells[10].text
                d_LUT = ourCells[11].text
                d_Att = ourCells[12].text
                
                print('{}-{} ({}) [{}]'.format(d_Forvalt, d_Todo, d_System, d_Rettighed))
                
                if d_Rettighed == 'AD bruger':
                    ActionChains(driver).move_to_element(r).double_click().perform()
                    print('Opretter AD bruger...')

                    sleep(5)
                    driver.find_element_by_id('ui-id-8').click()

                    sys.exit()


            ActionChains(driver).move_to_element(driver.find_element_by_id('div_main_content_back')).click().perform()
            # sleep(1)
            print('')
            

        slutTid = time()
        brugtTid = slutTid - startTid
        print('Dette loop tog {:.2f} sekunder. Genstarter om 2 minutter...'.format(brugtTid))

        sleep(120)
        os.system('cls')
        
    except Exception as e:
        print(e)
        sys.exit()
    
    except KeyboardInterrupt:
        sys.exit()

