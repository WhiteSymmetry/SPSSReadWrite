"""
SPSS.py

A Python module for importing SPSS files

(c) Alan James Salmoni
Released under the Affero General Public License


Notes: This only imports types 7 subtypes 3, 4, 5 and 6. Other subtypes are:
7: Multiple response set definitions
8: Data Entry for Windows (DEW) information
10: TextSmart information
11: Measurement level, column width and alignment for each variable - DONE!
13:
14:
17: text field defining variable attributes
20: Single string character encoding
21: Encodes value labels for long string variables

USAGE:

call SPSS.SPSSFile(args)

args are:

-all: to immediately import the file without waiting for commands to open
and read it
-pickle: to return the SPSS file pickled as a Python object (string format)
-help: to print this out
"""


import struct
import sys
import pickle


def pkint(vv):
    """
    An auxilliary function that returns an integer from a 4-byte word.
    The integer is packed in a tuple.
    """
    try:
        return struct.unpack("i",vv)
    except: # what is the error?
        return 0

def pkflt(vv):
    """
    An auxilliary function returns a double-precision float from an 8-byte word
    The float is packed in a tuple.
    """
    try:
        return struct.unpack("d",vv)
    except:
        return 0.0

def pkstr(vv):
    """
    An auxilliary function that returns a string from an 8-byte word. The 
    string is NOT packed.
    """
    bstr = ''
    for i in str(vv):
        bstr = bstr + struct.unpack("s",i)[0]
    return bstr


class variable(object):
    """
    This class contains a variable and its attributes. Each variable within 
    the SPSS file causes an instantiation of this class. The file object 
    contains a list of these in self.variablelist.
    """
    def __init__(self):
        self.name = None # 8 char limit
        self.namelabel = None
        self.data = []
        self.missingmarker = None
        self.missingd = []
        self.missingr = []
        self.type = None # 0 = numeric, 1 = string, -1 = string continuation
        self.printformatcode = []
        self.writeformatcode = []
        self.labelvalues = []
        self.labelfields = []

class SPSSFile(object):
    def __init__(self, *args):
        self.filename = args[0]
        self.fin = None
        self.typecode = []
        self.labelmarker = []
        self.missingvals = []
        self.missingvalmins = []
        self.missingvalmaxs = []
        self.documents = ''
        self.variablelist = []
        self.rawvarlist = []
        self.numvars = 0
        self.variablesets = None
        self.datevars = []
        if '-all' in args:
            self.OpenFile()
            self.GetRecords()

    def OpenFile(self):
        """
        This method trys to open the SPSS file.
        """
        try:
            self.fin = open(self.filename, "rb")
        except IOError:
            print "Cannot open file"
            self.fin = None

    def GetRecords(self):
        """
        This method read a 4-byte word and works out what record type it is, 
        and then despatches to the correct method. This continues until the 
        '999' code is reached (and of dictionary) upon which the data are read.
        """
        self.GetRecordType1()
        while 1:
            IN = pkint(self.fin.read(4))[0]
            if IN == 2:
                # get record type 2
                self.GetRecordType2()
            elif IN == 3:
                # get record type 3
                self.GetRecordType3()
            elif IN == 6:
                # get record type 6
                pass
            elif IN == 7:
                # get record type 7
                self.GetRecordType7()
            elif IN == 999:
                # last record end
                self.fin.read(4)
                self.GetData()
                self.fin.close()
                self.fin = None #need to remove file object for pickling
                return
        return

    def GetRecordType1(self):
        """
        This method reads in a type 1 record (file meta-data).
        """
        self.recordtype = self.fin.read(4)
        self.eyecatcher = self.fin.read(60)
        self.filelayoutcode = pkint(self.fin.read(4))
        self.numOBSelements = pkint(self.fin.read(4))
        self.compressionswitch = pkint(self.fin.read(4))
        self.caseweightvar = pkint(self.fin.read(4))
        self.numcases = pkint(self.fin.read(4))
        self.compressionbias = (self.fin.read(8))
        self.metastr = self.fin.read(84)

    def GetRecordType2(self):
        """
        This method reads in a type 2 record (variable meta-data).
        """
        x = variable()
        IN = pkint(self.fin.read(4))[0]
        x.typecode = IN
        if x == 0:
            x.type = "Numeric"
        else:
            x.type = "String"
        if x.typecode != -1:
            IN = pkint(self.fin.read(4))[0]
            x.labelmarker = IN
            IN = pkint(self.fin.read(4))[0]
            x.missingmarker = IN
            IN = self.fin.read(4)
            x.decplaces = ord(IN[0])
            x.colwidth = ord(IN[1])
            x.formattype = self.GetPrintWriteCode(ord(IN[2]))
            IN = self.fin.read(4)
            x.decplaces_wrt = ord(IN[0])
            x.colwidth_wrt = ord(IN[1])
            x.formattype_wrt = self.GetPrintWriteCode(ord(IN[2]))
            IN = pkstr(self.fin.read(8))
            nameblankflag = True
            x.name = IN
            for i in x.name:
                if ord(i) != 32:
                    nameblankflag = False
            if x.labelmarker == 1:
                IN = pkint(self.fin.read(4))[0]
                x.labellength = IN
                if (IN % 4) != 0:
                    IN = IN + 4 - (IN % 4)
                IN = pkstr(self.fin.read(IN))
                x.label = IN
            else:
                x.label = ''
            for i in range(abs(x.missingmarker)):
                self.fin.read(8)
            if x.missingmarker == 0:
                # no missing values
                x.missingd = None
                x.missingr = (None,None)
            elif (x.missingmarker == -2) or (x.missingmarker == -3):
                # range of missing values
                val1 = pkflt(self.fin.read(8))[0]
                val2 = pkflt(self.fin.read(8))[0]
                x.missingr = (val1, val2)
                if x.missingmarker == -3:
                    IN = pkflt(self.fin.read(8))[0]
                    x.missingd = IN
                else:
                    x.missings = None
            elif (x.missingmarker > 0) and (x.missingmarker < 4):
                # n(mval) missing vals
                tmpmiss = []
                for i in range(x.missingmarker):
                    IN = pkflt(self.fin.read(8))[0]
                    tmpmiss.append(IN)
                x.missingd = tmpmiss
                x.missingr = None
            if not nameblankflag:
                self.variablelist.append(x)
                self.rawvarlist.append(len(self.variablelist))
        elif x.typecode == -1:
            # read the rest
            try:
                self.rawvarlist.append(self.rawvarlist[-1])
            except:
                self.rawvarlist.append(None)
            self.fin.read(24)

    def GetRecordType3(self):
        """
        This method reads in a type 3 and a type 4 record. These always occur 
        together. Type 3 is a value label record (value-field pairs for 
        labels), and type 4 is the variable index record (which variables 
        have these value-field pairs).
        """
        # now record type 3
        self.r3values = []
        self.r3labels = []
        IN = pkint(self.fin.read(4))[0]
        values = []
        fields = []
        for labels in range(IN):
            IN = self.fin.read(8)
            IN = pkflt(IN)[0]
            values.append(IN)
            l = ord(self.fin.read(1))
            if (l % 8) != 0:
                l = l + 8 - (l % 8)
            IN = pkstr(self.fin.read(l-1))
            fields.append(IN)
        # get record type 4
        t = pkint(self.fin.read(4))[0]
        if t == 4:
            numvars = pkint(self.fin.read(4))[0]
            # IN is number of variables
            labelinds = []
            for i in range(numvars):
                IN = pkint(self.fin.read(4))[0]
                # this is index, store it
                labelinds.append(IN)
            for i in labelinds:
                ind = self.rawvarlist[i-1]
                self.variablelist[ind-1].labelvalues = values
                self.variablelist[ind-1].labelfields = fields
        else:
            print "Invalid subtype (%s)"%t
            return
            #sys.exit(1)

    def GetRecordType6(self):
        """
        This method retrieves the document record. 
        """
        # document record, only one allowed
        IN = pkint(self.fin.read(4))[0]
        self.documents = pkstr(self.fin.read(80*IN))

    def GetRecordType7(self):
        """
        This method is called when a type 7 record is encountered. The 
        subtype is then worked out and despatched to the proper method. Any 
        subtypes that are not yet programmed are read in and skipped over, so 
        not all subtype methods are yet functional.
        """
        # get subtype code
        subtype = pkint(self.fin.read(4))[0]
        if subtype == 3:
            self.GetType73()
        elif subtype == 4:
            self.GetType74()
        elif subtype == 5:
            self.GetType75()
        elif subtype == 6:
            self.GetType76()
        elif subtype == 11:
            self.GetType711()
        elif subtype == 13:
            self.GetType713()
        else:
            self.GetType7other()

    def GetType73(self):
        """
        This method retrieves records of type 7, subtype 3. This is for 
        release and machine specific integer information (eg, release 
        number, floating-point representation, compression scheme code etc).
        """
        # this is for release and machine-specific information
        FPrep = ["IEEE","IBM 370", "DEC VAX E"]
        endian = ["Big-endian","Little-endian"]
        charrep = ["EBCDIC","7-bit ASCII","8-bit ASCII","DEC Kanji"]
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0]
        if numelements == 8:
            self.releasenum = pkint(self.fin.read(4))[0]
            self.releasesubnum = pkint(self.fin.read(4))[0]
            self.releaseidnum = pkint(self.fin.read(4))[0]
            self.machinecode = pkint(self.fin.read(4))[0]
            self.FPrep = FPrep[pkint(self.fin.read(4))[0] - 1]
            self.compressionscheme = pkint(self.fin.read(4))[0]
            self.endiancode = endian[pkint(self.fin.read(4))[0] - 1]
            self.charrepcode = charrep[pkint(self.fin.read(4))[0] - 1]
        else:
            print "Error reading type 7/3"
            return
            #sys.exit(1)

    def GetType74(self):
        """
        This method retrieves records of type 7, subtype 4. This is for 
        release and machine-specific OBS-type information (system missing 
        value [self.SYSMIS], and highest and lowest missing values.
        """
        # release & machine specific OBS information
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0]
        if (numelements == 3) and (datatype == 8):
            self.SYSMIS = pkflt(self.fin.readline(8))[0]
            self.himissingval = pkflt(self.fin.readline(8))[0]
            self.lomissingval = pkflt(self.fin.readline(8))[0]
        else:
            print "Error reading type 7/4"
            return
            #sys.exit(1)

    def GetType75(self):
        """
        This method parses variable sets information. This is not 
        functional yet.
        """
        # variable sets information
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0]
        self.variablesets = pkstr(self.fin.read(4 * numelements))

    def GetType76(self):
        """
        This method parses TRENDS data variable information. This is not 
        functional yet.
        """
        # TRENDS data variable information
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0]
        # get data array
        self.explicitperiodflag = pkint(self.fin.read(4))[0]
        self.period = pkint(self.fin.read(4))[0]
        self.numdatevars = pkint(self.fin.read(4))[0]
        self.lowestincr = pkint(self.fin.read(4))[0]
        self.higheststart = pkint(self.fin.read(4))[0]
        self.datevarsmarker = pkint(self.fin.read(4))[0]
        for i in xrange(1, self.numdatevars + 1):
            recd = []
            recd.append(pkint(self.fin.read(4))[0])
            recd.append(pkint(self.GetDateVar(self.fin.read(4))[0]))
            recd.append(pkint(self.fin.read(4))[0])
            self.datevars.append(recd)

    def GetType711(self):
        """
        This method retrieves information about the measurement level, column 
        width and alignment.
        """
        measure = ["Nominal", "Ordinal", "Continuous"]
        align = ["Left", "Right", "Centre"]
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0] / 3
        for ind in range(numelements):
            var = self.variablelist[ind]
            IN = pkint(self.fin.read(datatype))[0]
            var.measure = measure[IN - 1]
            IN = pkint(self.fin.read(datatype))[0]
            var.displaywidth = IN
            IN = pkint(self.fin.read(datatype))[0]
            var.align = align[IN]

    def GetType713(self):
        """
        This method retrieves information about the long variable names 
        record. 
        """
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0]
        IN = self.fin.read(numelements)
        key = ''
        value = ''
        word = key
        for byte in IN:
            if ord(byte) == "=":
                word = value
            elif byte == '09':
                word = key
            else:
                word = word + byte

    def GetType7other(self):
        """
        This method is called when other subtypes not catered for are 
        encountered. See the introdoction to this module for more 
        information about their contents.
        """
        datatype = pkint(self.fin.read(4))[0]
        numelements = pkint(self.fin.read(4))[0]
        self.Other7 = self.fin.read(datatype * numelements)

    def GetData(self):
        """
        This method retrieves the actual data and stores them into the 
        appropriate variable's 'data' attribute.
        """
        self.cluster = []
        for case in range(self.numcases[0]):
            for i, var in enumerate(self.variablelist):
                if var.typecode == 0: # numeric variable
                    N = self.GetNumber()
                    if N == "False":
                        print "Error returning case %s, var %s"%(case, i)
                        sys.exit(1)
                    var.data.append(N)
                elif (var.typecode > 0) and (var.typecode < 256):
                    S = self.GetString(var)
                    if S == "False":
                        print "Error returning case %s, var %s"%(case, i)
                        sys.exit(1)
                    var.data.append(S)

    def GetNumber(self):
        """
        This method is called when a number / numeric datum is to be 
        retrieved. This method returns "False" (the string, not the Boolean 
        because of conflicts when 0 is returned) if the operation is not 
        possible.
        """
        if self.compressionswitch == 0: # uncompressed number
            IN = self.fin.read(8)
            if len(IN) < 1:
                return "False"
            else:
                return pkflt(IN)[0]
        else: # compressed number
            if len(self.cluster) == 0: # read new bytecodes
                IN = self.fin.read(8)
                for byte in IN:
                    self.cluster.append(ord(byte))
            byte = self.cluster.pop(0)
            if (byte > 1) and (byte < 252):
                return byte - 100
            elif byte == 252:
                return "False"
            elif byte == 253:
                IN = self.fin.read(8)
                if len(IN) < 1:
                    return "False"
                else:
                    return pkflt(IN)[0]
            elif byte == 254:
                return 0.0
            elif byte == 255:
                return self.SYSMIS

    def GetString(self, var):
        """
        This method is called when a string is to be retrieved. Strings can be 
        longer than 8-bytes long if so indicated. This method returns SYSMIS 
        (the string not the Boolean) is returned due to conflicts. 
        """
        if self.compressionswitch == 0:
            IN = self.fin.read(8)
            if len(IN) < 1:
                return self.SYSMIS
            else:
                return pkstr(IN)
        else:
            ln = ''
            while 1:
                if len(self.cluster) == 0:
                    IN = self.fin.read(8)
                    for byte in IN:
                        self.cluster.append(ord(byte))
                byte = self.cluster.pop(0)
                if (byte > 0) and (byte < 252):
                    return byte - 100
                if byte == 252:
                    return self.SYSMIS
                if byte == 253:
                    IN = self.fin.read(8)
                    if len(IN) < 1:
                        return self.SYSMIS
                    else:
                        ln = ln + pkstr(IN)
                        if len(ln) > var.typecode:
                            return ln
                if byte == 254:
                    if ln != '':
                        return ln
                if byte == 255:
                    return self.SYSMIS

    def GetPrintWriteCode(self, code):
        """
        This method returns the print / write format code of a variable. The 
        returned value is a tuple consisting of the format abbreviation 
        (string <= 8 chars) and a meaning (long string). Non-existent codes 
        have a (None, None) tuple returned.
        """
        if type(code) != int:
            return
        if code == 0:
            return ('','Continuation of string variable')
        elif code == 1:
            return ('A','Alphanumeric')
        elif code == 2:
            return ('AHEX', 'alphanumeric hexadecimal')
        elif code == 3:
            return ('COMMA', 'F format with commas')
        elif code == 4:
            return ('DOLLAR', 'Commas and floating point dollar sign')
        elif code == 5:
            return ('F', 'F (default numeric) format')
        elif code == 6:
            return ('IB', 'Integer binary')
        elif code == 7:
            return ('PIBHEX', 'Positive binary integer - hexadecimal')
        elif code == 8:
            return ('P', 'Packed decimal')
        elif code == 9:
            return ('PIB', 'Positive integer binary (Unsigned)')
        elif code == 10:
            return ('PK', 'Positive packed decimal (Unsigned)')
        elif code == 11:
            return ('RB', 'Floating point binary')
        elif code == 12:
            return ('RBHEX', 'Floating point binary - hexadecimal')
        elif code == 15:
            return ('Z', 'Zoned decimal')
        elif code == 16:
            return ('N', 'N format - unsigned with leading zeros')
        elif code == 17:
            return ('E', 'E format - with explicit power of ten')
        elif code == 20:
            return ('DATE', 'Date format dd-mmm-yyyy')
        elif code == 21:
            return ('TIME', 'Time format hh:mm:ss.s')
        elif code == 22:
            return ('DATETIME', 'Date and time')
        elif code == 23:
            return ('ADATE', 'Date in mm/dd/yyyy form')
        elif code == 24:
            return ('JDATE', 'Julian date - yyyyddd')
        elif code == 25:
            return ('DTIME', 'Date-time dd hh:mm:ss.s')
        elif code == 26:
            return ('WKDAY', 'Day of the week')
        elif code == 27:
            return ('MONTH', 'Month')
        elif code == 28:
            return ('MOYR', 'mmm yyyy')
        elif code == 29:
            return ('QYR', 'q Q yyyy')
        elif code == 30:
            return ('WKYR', 'ww WK yyyy')
        elif code == 31:
            return ('PCT', 'Percent - F followed by "%"')
        elif code == 32:
            return ('DOT', 'Like COMMA, switching dot for comma')
        elif (code >= 33) and (code <= 37):
            return ('CCA-CCE', 'User-programmable currency format')
        elif code == 38:
            return ('EDATE', 'Date in dd.mm.yyyy style')
        elif code == 39:
            return ('SDATE', 'Date in yyyy/mm/dd style')
        else:
            return (None, None)

    def GetDateVar(self, code):
        datetypes = [   "Cycle","Year","Quarter","Month","Week","Day","Hour",
                        "Minute","Second","Observation","DATE_"]
        try:
            return datetypes[code]
        except IndexError:
            return None

    def GetNames(self):
        """
        This method retrieves all the names for all the variables. If the file has
        not already been read, strange results (though probably a blank) will be
        returned.
        """
        names = []
        for variable in self.variablelist:
            names.append(variable.name)
        return names

    def GetLabels(self):
        """
        This method returns the labels of all the variables
        """
        labels = []
        for variable in self.variablelist:
            labels.append(variable.label)
        return labels

    def GetTypeCodes(self):
        """
        This method returns the typecodes of each variable.
        """
        typecodes = []
        for variable in self.variablelist:
            typecodes.append(variable.typecode)
        return typecodes

    def GetRow(self, row):
        """
        This method returns a row of data
        """
        if (row < 0) or (row > self.numcases):
            return None
        else:
            row = []
            for ind, variable in enumerate(self.variablelist):
                row.append(variable.data[ind])
            return row


if __name__ == '__main__':
    args = sys.argv
    args.pop(0)
    x = SPSSFile(args)
    if "-pickle" in args:
        p = pickle.dumps(x)
        print p
    if "-help" in args:
        print "SPSS file importer for Python"
        print "(c) 2008 Alan James Salmoni [salmoni at gmail]"
        print "Command line arguments:"
        print "SPSS.SPSSFile file args"
        print "file is valid file name of SPSS (.sav) file"
        print "Args:"
        print "-all: immediately open and import the file"
        print "-pickle: return the SPSS file as a pickled Python object (string)"
        print "-help: print this"
    # How to use this
    f = 'C:\Documents and Settings\Authorized User\My Documents\Documents\AQ.sav'
    #x = SPSSFile(f)
    #x.OpenFile()
    #x.GetRecords()
    # then use commands like this to access data & metadata

    # FILE META-DATA:
    # x.eyecatcher # shows OS, machine, SPSS version etc
    # x.numOBSelements # puted number of variables (use x.numvars instead)
    # x.compressionswitch # 0 if not compressed
    # x.metastr # creation date, time, file label.
    # x.variablelist # list of variable objects contained within
    # x.numvars # number of variables.
    # x.documents # documentation record (if any)

    # VARIABLE META-DATA:
    # y = x.variablelist
    # y.data # the data
    # y.name # 8-byte variable name
    # y.label # longer string label
    # y.decplaces # number of decimal places
    # y.colwidth # column width
    # y.formattype # print format code (the exact data type)
    # y.labelvalues # values for substitute labels
    # y.labelfields # fields for substitute labels
    # y.missingd # list of discrete missing values
    # y.missingr # upper and lower bounds of a range of missing values
    # and many more.
    # Check dir: lower case starts are attributes, others are methods

    # EXTRAS:

    # * GetNames method to return names from all variables
    # * GetRows method to return data from particular row
    # * GetLabels method to return labels from all variables
    # * GetTypeCodes method to return variables' typecodes

    # NEED TO ADD:

    # * working methods for various type 7 subtypes (all meta-data, some documented, not all)
    # * Any others?
