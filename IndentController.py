import pandas as pd
import time
import serial

def sendSerial(contents, comPort):
    # Send format = "STRING" + CR
    CR = 0x0d

    # Open port
    setSerial = serial.Serial(comPort, 19200)
    print(contents)

    # V0:partno, V1:lotno, V2:idno
    for i in range(len(contents)):
        sendCommand = []
        for sentence in contents[i]:
            sendCommand.append(ord(sentence))

        # sendCommand.append(0x0d)
        sendCommand.append(CR)
        
        # Convert binary
        sendCommand = bytearray(sendCommand)

        # Send serial code
        setSerial.write(sendCommand)
        time.sleep(0.1)

    setSerial.close()
    print("Finish sending text.")