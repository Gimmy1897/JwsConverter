#Gimmy1897.Dev();

import sys
import olefile as ofio
from struct import unpack
from struct import calcsize
import numpy as np
import os
import binascii
from os import path


#from jwsprocessor 
DATAINFO_FMT = '<LLLLLLdddLLLLdddd'
#from jwsprocessor

#header dictionary
headers_dict ={
    "00000000": "%Transmittance",
    "02000000": "%Reflectance",
    "03000000": "Absorbance",
    "06000000": "KM",
    "0E000000": "Intensity",
    "0F000000": "n",
    "10000000": "k",
    "01100000": "CD(mdeg)",
    "02100000": "ORD(mdeg)",
    "19100000": "FDCD(mdeg)",
    "20100000": "CD/DC(mdeg)",
    "21100000": "FDCD/DC(mdeg)",
    "22100000": "Test(mdeg)",
    "23100000": "LB(mdeg)",
    "24100000": "CB(mdeg)",
    "03100000": "LD(dOD)",
    "25100000": "FDLD(dOD)",
    "28100000": "LD/DC(dOD)",
    "06100000": "DC(V)",
    "01200000": "HT(V)",
    "04200000": "External(V)",
    "07100000": "Mol.Ellip.",
    "08100000": "Mol.CD",
    "09100000": "Spc.Ellip.",
    "16100000": "DeltaOD",
    "0A100000": "Mol.Rotation",
    "0B100000": "Spc.Rotation",
    "0C100000": "Spc.Abs",
    "0D100000": "Mol.Abs",
    "0E100000": "Mol.Ellip. MG",
    "0F100000": "Mol.CD MG",
    "10100000": "Spc.Ellip. MG",
    "11100000": "Mol.Rot MG",
    "12100000": "Spc.Rot MG",
    "13100000": "LDcor",
    "14100000": "DeltaEps.(LD)",
    "15100000": "DeltaE(LD)",
    "17100000": "A.U.",
    "1D000000": "G Value",
    "26100000": "g(lum)",
    "27100000": "Delta I",
    "24000000": "dAbs",
    "28000000": "Epsilon",
    "29000000": "DeltaEps.",
    "30000000": "dInt.",
    "11000000": "Delta",
    "12000000": "Psi",
    "13000000": "50KHz",
    "14000000": "100KHz",
    "15000000": "Retardation",
    "16000000": "Ellipsity",
    "17000000": "Rotation",
    "2000000": "P",
    "21000000": "r",
    "00010010": "Wavenumber(cm-1)",
    "01010010": "Raman Shift(cm-1)",
    "02010010": "Angstrom",
    "03010010": "Wavelength(nm)",
    "04010010": "Wavelength(um)",
    "09010010": "Point",
    "0a010210": "Channel",
    "00040010": "Energy[eV]",
    "01020020": "usec",
    "02020020": "msec",
    "03020020": "sec",
    "04020020": "min",
    "05020020": "hour",
    "00030030": "Kelvin",
    "01030030": "Celcius"
}


class DataInfo:
    def __init__(self, data):
        
        if len(data) < 96:
            raise Exception ("[JwsConverter Error]: this .jws file is not supported, DataInfo should be at least 96 bytes!")

        data = data[:96]
        data_tuple = unpack(DATAINFO_FMT, data)

        self.nchannels = data_tuple[3]
        self.npoints = data_tuple[5]
        self.x_for_first_point = data_tuple[6]
        self.x_for_last_point = data_tuple[7]
        #if x_increment is not 0, set it to the value from the file
        if(data_tuple[8]!=0):
            self.x_increment = data_tuple[8]
        else:
            #if x_increment is 0, calculate it
            self.x_increment = (self.x_for_last_point - self.x_for_first_point) / (self.npoints - 1)

        #get hex values of data
        hex_list = binascii.hexlify(data).decode().upper()
        #cut off the first 96 hex values (data_info)
        hex_list = hex_list[96:]
        #set X header, first 8 hex values
        self.x_header = headers_dict.get(hex_list[:8],"X")
        #cut off the first 8 hex values (x_header)
        hex_list = hex_list[8:]
        #set Y headers for each channel. 8 hex values for each channel
        y_headers = []
        for i in range(0,self.nchannels):
            #get header for each channel. If not found, set it to Y1, Y2, Y3, etc.
            y_header = headers_dict.get(hex_list[8*i:8*(i+1)],f"Y{i+1}")
            y_headers.append(y_header)
        self.y_headers = y_headers

def main(fn, output=None):
    xdata =[]
    ydata =[]
    nchannel =0
    #open file
    doc = ofio.OleFileIO(fn)
    #check if file has DataInfo section
    if doc.exists('DataInfo'):
        str = doc.openstream('DataInfo')
        data = str.read()
        DI = DataInfo(data)   
        nchannel = DI.nchannels
    else:
        print("[JwsConverter ERROR]: Oops! This is not .jws file, i can't convert it :((")
        return
    #check if file has Y-Data section
    if doc.exists('Y-Data'):
            #opens (Y-Data string(
            str = doc.openstream('Y-Data')
            data = str.read()
            #set fmt for unpacking data
            fmt = 'f' * DI.npoints
            #check if y data have the same length as the number of points
            if(len(data)/nchannel==calcsize(fmt)):
                #split data into nchannel parts
                for i in range(0,nchannel):
                    data_1 = data[int(i*len(data)/nchannel):int((i+1)*len(data)/nchannel)]
                    #get values for Yi in tuple form -- deciphers data string using struct
                    values = unpack(fmt, data_1)
                    ydata.append(list(values))
            else: 
                print("[JwsConverter ERROR]: Oops! This file is not properly formatted, i can't convert it :((")
                return

            #generate xdata
            x = DI.x_for_first_point
            x += DI.x_increment
            xdata = np.arange(DI.x_for_first_point,  # start
                                DI.x_for_last_point + DI.x_increment,  # end+incr.
                                DI.x_increment)  #
            #save data to txt file
            #get path of the file
            jwspath = file_path.rsplit("\\",1)
            #get directory path
            directory = os.path.dirname(jwspath[0]+'\\txt\\')
            #check if directory exists, if not create it
            if not os.path.exists(directory):
                os.makedirs(directory)
            #save data to txt file
            np.savetxt(jwspath[0]+'\\txt\\'+jwspath[1]+'.txt', np.column_stack((xdata, *ydata)), fmt='%1.8f',delimiter='\t', header=DI.x_header+'\t'+'\t'.join(DI.y_headers))

    else:
        print("[JwsConverter ERROR]: Oops! This is not .jws file, i can't convert it :((")
        return
if __name__ == '__main__':
    print("[JwsConverter]: Welcome to JwsConverter :P !\n______\n")
    if len(sys.argv) > 1:
        paths = sys.argv[1]

        for file_path in paths.split(","):
            try:
                dir_path = file_path.rsplit("\\",1)
            except: 
                print("\n[JwsConverter]:")

            if(path.exists(dir_path[0])):

                os.chdir(dir_path[0])

                print("[JwsConverter]: Conversion of ----> ",dir_path[1])
                main(dir_path[1])
                print("[JwsConverter]: Done :))\n______\n")
                  
            else:   
                print ("[JwsConverter ERROR]: The selected path (",file_path,") is unreachable :/  change it!") 
    else:
        print ('USAGE:py jwsconverter_smf.py input_file_paths')



# Gimmy1897.Dev();
