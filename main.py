import pyautogui as pg
import time
from datetime import datetime
import logging
import clipboard as copy
import pandas as pd
import atexit


# 기본 값 정의

processSet = 'Data File Name'

dat = pd.read_csv('./{}.csv'.format(processSet), encoding='utf-8')
datLen = len(dat)

## Image Icon Path
homePath = './image/Home/'
exportPath = './image/Export/'
exportListPath = './image/ExportList/'
commonPath = './image/'

## Log Path
logPath = './log/'
logScreenShotPath = './log/screenshot/'

## result Path
resultPath = './result/result/'
resultErrorPath = './result/error/'
resultInterruptPath = './result/inter/'

## export List status 정의 List
statusList = ['Canceled Export', 'Canceling Export','Complete Export', 'Error Export', 'Error Privilege', 'Start Export']

## findImg confidence 정의

imgConfidence = .9

## position x, y 이동 값 정의 dictionary : list
posDicHome = {'export': [60, 40], 'exportList': [60, 60]}
posDicExportList = {'fstStatus': [0, 60], 'popStatus': [1140,516], 'popTotalCnt': [1140,555], 'dragLength':[-250, 0], 'copyPos':[30,3]}



## 함수 정의

### 프로그램 있는지 확인

def activeProgram(logger):

    win = pg.getWindowsWithTitle('Program Title')
    if(win is not None):
        return win
    else:
        logger.error('program이 켜지지 않았습니다.')
        return 1;

### Image box 객체 반환 함수

def findImg(path, accurate): #return boxObj
    findingCnt = 0
    img = pg.locateOnScreen(path, confidence=accurate)

    while(img is None):
        findingCnt+=1
        print('findingImg_Cnt:' + str(findingCnt))
        if(findingCnt > 9):
            break
        time.sleep(1)
        img = pg.locateOnScreen(path, confidence=accurate)


    if(img is None):
        return 1;
    else:
        pg.moveTo(img.left, img.top)
        return pg.position()

def findImgCenter(path, accurate):
    findingCnt = 0
    img = pg.locateCenterOnScreen(path, confidence=accurate)

    while(img is None):
        findingCnt+= 1
        print('findingImgCenter_Cnt:'+ str(findingCnt))
        if(findingCnt > 9):
            break
        time.sleep(1)
        img = pg.locateCenterOnScreen(path, confidence=accurate)

    if (img is None):
        return 1
    else:
        pg.moveTo(img.x, img.y)
        return pg.position()

def setPosition(pos, X, Y):
    try:
        pg.moveTo(pos.x+X, pos.y+Y)
        return pg.position()
    except Exception as e:
        pg.alert(e,"경고")
        return 1

def returnHome(logger):
    returnHome = findImgCenter(commonPath+'returnHome.png',imgConfidence)

    if(returnHome != 1):
        setPosition(returnHome, 80,0)
        pg.click()
        time.sleep(3)
        return 0
    else:
        logger.error('return Home 아이콘을 찾을 수 없음')
        return 1

### check Home

def checkHome(logger):
    home = findImg(homePath +'homeCheck.png',imgConfidence)
    if(home == 1):
        returnHome(logger)
        home = findImg(homePath + 'homeCheck.png', imgConfidence)
    if(home != 1):
        return 0
    else:
        logger.error('program Home 화면이 아님')
        return 1

### Export 창 이동

def moveExport(logger):
    if(checkHome(logger) != 1):
        exportRegion = findImg(homePath + 'exportRegion.png',imgConfidence)
        setPosition(exportRegion, posDicHome['export'][0], posDicHome['export'][1])
        pg.click()
        time.sleep(5)
        return 0
    else:

        return 1

### ExportList 창 이동
def moveExportList(logger):
    if(checkHome(logger) != 1):
        exportRegion = findImg(homePath + 'exportRegion.png', imgConfidence)
        setPosition(exportRegion, posDicHome['exportList'][0], posDicHome['exportList'][1])
        pg.click()
        time.sleep(3)
        return 0
    else:
        logger.error('현재 홈 화면이 아닙니다.')
        return 1

def strToCnt(currentTotal):
     ct = currentTotal.replace('(Completed ', '')
     ct = ct.rstrip(')')
     [totalCnt, curCnt] = ct.split()

     return [int(totalCnt), int(curCnt)]


#### 자동 연결 끊김 방지를 위해 시간 조정
def sleepTimer(cntList):
    idx = ((cntList[0] - cntList[1]) // 2000 + 1 )
    print('기다리는 시간(분): {}'.format(idx))
    if (idx>15):
        idx = 15
        print('조정된 기다리는 시간(분): {}'.format(idx))
    while(idx>0):
        time.sleep(60)
        idx -=1
    return 0


#### ExportList 현재 status 확인

def currentStatus(logger):
    statusIconPath = exportListPath +'statusCol.png'
    popStatusPath = exportListPath + 'popStatus.png'
    popTotalPath = exportListPath + 'popTotal.png'
    popOkPath = exportListPath + 'popOk.png'
    statusIcon = findImgCenter(statusIconPath, imgConfidence)

    if(statusIcon != 1):
        # logger.info('Export List에서 첫번째 Status Line 찾음')
        setPosition(statusIcon, posDicExportList['fstStatus'][0], posDicExportList['fstStatus'][1])
        pg.click()
        time.sleep(1)
        popStatus =findImgCenter(popStatusPath, imgConfidence)
        if(popStatus == 1):
            logger.error('Export Status를 찾을 수 없습니다.')
            return 1
        # else:
            # logger.info('Export Status 찾음')

        pg.moveTo(posDicExportList['popStatus'][0],posDicExportList['popStatus'][1])
        time.sleep(1)

        # 복사 기능 - Status

        # pg.doubleClick()
        # pg.hotkey('ctrl', 'a')
        # time.sleep(1)
        # pg.hotkey('ctrl', 'c')
        # time.sleep(1)

        pg.drag(posDicExportList['dragLength'][0],posDicExportList['dragLength'][1])
        time.sleep(1)
        pg.move(posDicExportList['copyPos'][0],posDicExportList['copyPos'][1])
        pg.click(button='right')
        time.sleep(1)
        pg.move(posDicExportList['copyPos'][0],posDicExportList['copyPos'][1])
        time.sleep(1)
        pg.click()
        currentStatus = copy.paste()
        print('현재 상태: '+ currentStatus)




        # 복사 기능 - Source Total
        popTotal = findImgCenter(popTotalPath, imgConfidence)
        if(popTotal ==1):
            logger.error('Export Total 찾을 수 없습니다.')
            return 1
        # else:
        #     logger.info('Export Total 찾음')
        # setPosition(popTotal, posDicExportList['popStatus'][0], posDicExportList['popStatus'][1])

        time.sleep(1)
        pg.moveTo(posDicExportList['popTotalCnt'][0],posDicExportList['popTotalCnt'][1])
        time.sleep(1)
        pg.drag(posDicExportList['dragLength'][0], posDicExportList['dragLength'][1])
        time.sleep(1)
        pg.move(posDicExportList['copyPos'][0], posDicExportList['copyPos'][1])
        pg.click(button='right')
        time.sleep(1)
        pg.move(posDicExportList['copyPos'][0], posDicExportList['copyPos'][1])
        time.sleep(1)
        pg.click()
        currentCount = copy.paste()
        print('현재 진행된 count: '+currentCount)


        # pg.doubleClick()
        # pg.hotkey('ctrl', 'a')
        # time.sleep(1)
        # pg.hotkey('ctrl', 'c')
        # time.sleep(1)
        # currentCount = copy.paste()
        # print('현재 진행된 count: '+currentCount)

        popOk = findImgCenter(popOkPath, imgConfidence)
        if(popOk ==1):
            logger.warning('Export List Status Pop-up OK버튼을 찾을 수 없습니다.')
            return 1
        # else:
        #     logger.info('Export List Status Pop-up OK버튼을 찾음')
        pg.click()

        return [currentStatus, currentCount]

    else:
        # statusIcon을 찾을 수 없습니다.
        logger.warning('Export List에서 첫번째 Status Line 찾을 수 없습니다.')
        return 1

def screenShot(targetId):

    try:
        SSTime = datetime.now().strftime('SS_%y-%m-%d_%H%M%S')
        path = logScreenShotPath + SSTime + '_' + targetId + '.png'
        currentImg = pg.screenshot()
        currentImg.save(path)
        return 0
    except Exception as e:
        pg.alert(e, '경고')
        return 1

def warn_stop(logger, target, targetName):
    if(target ==1):
        logger.warning('export창에서 {}을 찾을 수 없습니다.'.format(targetName))
        # logger.warning(targetName)
        pg.alert('{}을 찾을 수 없습니다.'.format(targetName), '정지')
    return 0



def exporting(targetId, exportFolderName, logger):
    logger.warning(targetId + '_exporting setting start')
    print(targetId+ '_exporting setting start')

    borderLinePath = exportPath + 'borderLine.png'
    exportButtonPath = exportPath + 'exportButton.png'
    exportDirInputPath = exportPath + 'exportDirInput.png'
    loadPresetPath = exportPath + 'loadPreset.png'
    presetItemPath = exportPath + 'presetItem_Prod.png'
    presetLoadButtonPath = exportPath + 'presetLoadButton.png'
    presetPublicRadioPath = exportPath + 'presetPublicRadio.png'
    # reviewSelectButtonPath = exportPath + 'reviewSelectButton.png'
    # selectCasePath = exportPath + 'selectCase.png'
    # selectCaseIconPath = exportPath + 'selectCaseIcon.png'
    targetIdIconPath = exportPath + 'targetIdIcon.png'
    targetIdInputPath = exportPath + 'targetIdInput.png'
    targetSelectRadioPath = exportPath + 'targetSelectRadio.png'
    afterExportOkPath = exportPath +'afterExportOk.png'

    waitLoading = 0

    while(waitLoading < 3):
        borderLine = findImg(borderLinePath, imgConfidence)
        warn_stop(logger, borderLine, 'borderLine')
        setPosition(borderLine, -5, 50)
        print(pg.position())
        pg.drag(600, 0)
        time.sleep(1)
        targetIdIcon = findImgCenter(targetIdIconPath, imgConfidence)
        if(targetIdIcon != 1):
            break
        waitLoading += 1

    warn_stop(logger, targetIdIcon, 'targetIdIcon')
    setPosition(targetIdIcon, 30, 0)
    pg.click()
    time.sleep(1)
    targetIdInput = findImgCenter(targetIdInputPath, imgConfidence)
    warn_stop(logger, targetIdInput, 'targetIdInput')
    setPosition(targetIdInput, 0, -15)
    pg.click()
    pg.write(targetId)
    setPosition(targetIdInput, 40, 40)
    pg.click()
    time.sleep(1)

    targetSelectRadio = findImgCenter(targetSelectRadioPath,imgConfidence)
    warn_stop(logger, targetSelectRadio, 'targetSelectRadio')
    setPosition(targetSelectRadio, -40, 10)
    pg.click()

    loadPreset = findImgCenter(loadPresetPath, imgConfidence)
    warn_stop(logger, loadPreset, 'loadPreset')
    pg.click()
    time.sleep(1)
    presetPublicRadio = findImgCenter(presetPublicRadioPath, imgConfidence)
    warn_stop(logger, presetPublicRadio, 'presetPublicRadio')
    pg.click()
    time.sleep(1)

    presetItem = findImgCenter(presetItemPath, imgConfidence)
    warn_stop(logger, presetItem, 'presetItem')
    pg.click()
    presetLoadButton = findImgCenter(presetLoadButtonPath, imgConfidence)
    warn_stop(logger, presetLoadButton, 'presetLoadButton')
    pg.click()

    exportDirInput = findImgCenter(exportDirInputPath, imgConfidence)
    warn_stop(logger, exportDirInput, 'exportDirInput')
    pg.doubleClick()
    pg.press('backspace')
    pg.hotkey('ctrl', 'a')
    #  원래는 targetName...
    copy.copy(exportFolderName)
    pg.rightClick()
    pg.move(5, 30)
    pg.click()

    exportButton = findImgCenter(exportButtonPath, imgConfidence)
    warn_stop(logger, exportButton, 'exportButton')
    screenShot(targetId)

    print(targetId+'_export를 시작합니다.')
    pg.countdown(10)
    exportButton = findImgCenter(exportButtonPath, imgConfidence)
    warn_stop(logger, exportButton, 'exportButton')
    pg.click()
    logger.warning(targetId + '_export를 시작합니다.')
    afterExportOk = findImgCenter(afterExportOkPath, imgConfidence)
    warn_stop(logger, afterExportOk, 'afterExportOk')
    pg.click()

    return 0

def saveWithError(dat):
    filename = datetime.now().strftime('{}_withError_%y-%m-%d_%H%M%S.xlsx'.format(processSet))
    filename = resultErrorPath + filename
    print(filename)
    dat.to_excel(filename, encoding='utf-8')
    return 0

def saveResult(dat):
    filename = datetime.now().strftime('{}_%y-%m-%d_%H%M%S.xlsx'.format(processSet))
    filename = resultPath + filename
    print(filename)
    dat.to_excel(filename, encoding='utf-8')

def saveWithInterrupt(dat):
    filename = datetime.now().strftime('{}_withInterrupt_%y-%m-%d_%H%M%S.xlsx'.format(processSet))
    filename = resultInterruptPath + filename
    print(filename)
    dat.to_excel(filename, encoding='utf-8')
    return 0

atexit.register(saveWithInterrupt, dat)

def main():

    try:

        logFormatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)

        streamHandler = logging.StreamHandler()
        streamHandler.setFormatter(logFormatter)
        logger.addHandler(streamHandler)

        filename = datetime.now().strftime('log_%y-%m-%d_%H%M.log')
        filename = logPath + filename
        fileHandler = logging.FileHandler(filename, encoding='utf-8')
        fileHandler.setFormatter(logFormatter)
        logger.addHandler(fileHandler)

        i = 0
        isMove = 0
        isHome = 0

        if(i==0):

            if (activeProgram(logger) == []):
                return 1
            isHome = checkHome(logger)
            isMove = moveExportList(logger)
            if ( isHome+isMove > 0 ):
                pg.alert('Export List 화면 진입 실패', '경고')

            currentState = currentStatus(logger)

            logger.info('시작 전 Export 상태: {}'.format(currentState[0]))
            logger.info('시작 전 Export count: {}'.format(currentState[1]))

            if (statusList.index(currentState[0]) == 5):
                # print(statusList.index(currentState[0]))
                while(statusList.index(currentState[0])==5):
                    sleepTimer(strToCnt(currentState[1]))
                    currentState = currentStatus(logger)

            isMove = returnHome(logger)
            isHome = checkHome(logger)

            if(isMove + isHome >0):
                pg.alert('현재 홈 화면이 아닙니다.', '경고')



        while(i < datLen):

            if(i != 0 and i%10 ==0):
                saveResult(dat)

            logger.debug('{}번째 data 시작'.format(i))
            print('{}번째 data 시작'.format(i))

            currentTargetId = dat['targetId'][i]
            if(pd.isnull(dat['exportStatus'][i]) == False):
                logger.warning('{}은 이미 export 완료 되었습니다.'.format(currentTargetId))
                i +=1
                continue

            isMove = moveExport(logger)
            if(isMove > 0):
                pg.alert('{}번째 data에서 Export 창 진입에 실패했습니다.', '경고')

            exporting(dat['targetId'][i], dat['exportFolderName'][i], logger)
            dat.loc[dat['targetId'] == currentTargetId, 'exportStartTime'] = datetime.now().strftime('%y-%m-%d_%H:%M:%S')
            isMove = returnHome(logger)
            isHome = checkHome(logger)
            isMove = moveExportList(logger)
            if(isMove + isHome > 0):
                pg.alert('{} export 후 exportList 화면 진입 실패'.format(dat['targetId'][i]))

            currentState = currentStatus(logger)

            while len(currentState[1]) < 5:
                time.sleep(5)
                currentState = currentStatus(logger)


            cntList = strToCnt(currentState[1])
            logger.info('현재 {} 상태: {}'.format(dat['targetId'][i], currentState[0]))
            logger.info('현재 {} count: {}'.format(dat['targetId'][i], currentState[1]))
            dat.loc[dat['targetId'] == currentTargetId, 'totalCnt'] = cntList[0]

            while(statusList.index(currentState[0]) != 2):
                if(statusList.index(currentState[0]) == 3):
                    result = pg.confirm('현재 {} Export에 Error가 발생했습니다. 중지하시겠습니까?'.format(currentTargetId), '확인')
                    if(result =='확인'):
                        saveWithError(dat)
                        return 1

                sleepTimer(strToCnt(currentState[1]))
                # time.sleep(dat['checkSec'][i])
                currentState = currentStatus(logger)
                logger.info('현재 {} 상태: {}'.format(dat['targetId'][i] , currentState[0]))
                logger.info('현재 {} count: {}'.format(dat['targetId'][i], currentState[1]))

            dat.loc[dat['targetId'] == currentTargetId, 'exportEndTime'] = datetime.now().strftime('%y-%m-%d_%H:%M:%S')
            dat.loc[dat['targetId'] == currentTargetId, 'exportStatus'] = currentState[0]
            logger.warning('{} exporting 완료'.format(dat['targetId'][i]))

            i += 1
            isMove = returnHome(logger)
            isHome = checkHome(logger)

            if(isHome + isMove > 0):
                pg.alert('현재 홈 화면 진입이 실패했습니다.', '경고')
        saveResult(dat)
        return 0
    except Exception as e:
        saveWithError(dat)
        pg.alert(e, '경고')
        return 1


main()
