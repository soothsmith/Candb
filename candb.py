
"""\
This module provides CAN database(*dbc) and matrix(*xls) operation functions.
-- Convert a specific format of spreadsheet into a DBC file.
-- Sort a DBC file by message/signal.
-- Merge multiple DBC files into a single output DBC file.
---   Content is loaded and may be overwritten by each 
---   successive file in the order they appear on the command line.
---   Hint - to ensure compatibility with the DBC file, 
---   use the merge function on a single file and compare 
---   input to output (attribute order is likely to be different).
-- Compare is not yet implemented.
-- 
Content is translated according to: 
-- DBC File Format Documentation, Version 01/2007
-- Example files for load/sort/write included SAE_j1939.dbc
-- To merge dbc files, node attributes, comments, environment variables,
---  value types have been added.
-- Also able to read and dump Environment varaibles and value tables.
-- Added ability to read comments spannding multiple lines (except for Nodes).
-- Added ability to read/store multiplexing values.
---  Added a MultiplexorIndicator column to the spreadsheet.
-- Now allow message comments (same column in spreadsheet as for signals).
--
Updates 18-Feb-2020
-- Improved the merge function for:
---- The way the Nodes are merged and how they store their attributes.
---- The way that attribute enumeration values are pulled in and stored so they reduce 
----    or eliminate duplicates for various upper/lower-case variants, and to ensure 
----    that they are encoded (for each input file) and decoded (from a combined list) correctly. 
---- The SG_MUL_VAL_ records are read and dumped verbatim since I have no other information about their definition.
---- VAL_ and VAL_TABLE_ entries (list of enumerations) loaded from files may now span multiple lines
-- Other info: 
---- If a single file is “merged” to output, the files will still be different because: 
----   1. More attributes are dumped (than are read from the files) as I’ve gathered them from various files and added them to the script.
----   2. Many of the message and signal attributes dump in a different order than how they are read (due to the way the attributes are stored).
"""
import re
import xlrd
import sys
#imort importlib
#import traceback
import os
#from builtins import False
#from builtins import True
#from scipy.stats.tests.test_discrete_basic import vals

#if sys.version_info[0] < 3:
#    reload(sys)
#    sys.setdefaultencoding('utf-8')
#else:
#    importlib.reload(sys)

# enable or disable debug info display, this switch is controlled by -d option.
debug_enable = False

# symbol definition of DBC format file
NEW_SYMBOLS = [
    'NS_DESC_', 'CM_', 'BA_DEF_', 'BA_', 'VAL_', 'CAT_DEF_', 'CAT_',
    'FILTER', 'BA_DEF_DEF_', 'EV_DATA_', 'ENVVAR_DATA_', 'SGTYPE_',
    'SGTYPE_VAL_', 'BA_DEF_SGTYPE_', 'BA_SGTYPE_', 'SIG_TYPE_REF_',
    'VAL_TABLE_', 'SIG_GROUP_', 'SIG_VALTYPE_', 'SIGTYPE_VALTYPE_',
    'BO_TX_BU_', 'BA_DEF_REL_', 'BA_REL_', 'BA_DEF_DEF_REL_',
    'BU_SG_REL_', 'BU_EV_REL_', 'BU_BO_REL_', 'SG_MUL_VAL_'
]

SignalAttributes_List = ["SPN", "DDI", "SigType", "GenSigInactiveValue", "GenSigSendType", "GenSigStartValue", "GenSigTimeoutValue", "SAEDocument", "SystemSignalLongSymbol", "GenSigILSupport", "GenSigCycleTime", "GenSigCycleTimeActive"]
### Signal value lists: 
Boolean_List         = ["No",    "Yes"]
Boolean_Order        = {'NO': 0, 'YES': 1}

NmNode_List          = ['Not',   'Yes']
NmNode_Order         = {'NOT': 0, 'YES': 1}

#GenMsgSendType_List  = ["Cyclic",   "cyclicX",   "spontanX",   "IfActive",   "cyclicIfActive",   "cyclicIfActiveX",   "spontanWithDelay",   "cyclicAndSpontanX",   "cyclicAndSpontanWithDelay",   "spontanWithRepetition",   "cyclicIfActiveAndSpontanWD",    "cyclicIfActiveFast",    "cyclicWithRepeatOnDemand",    "NoSigSendType",    "noMsgSendType",    "NotUsed",    "reserved"]
#GenMsgSendType_Order = {"CYCLIC": 0,"CYCLICX": 1,"SPONTANX": 2,"IFACTIVE": 3,"CYCLICIFACTIVE": 4,"CYCLICIFACTIVEX": 5,"SPONTANWITHDELAY": 6,"CYCLICANDSPONTANX": 7,"CYCLICANDSPONTANWITHDELAY": 8,"SPONTANWITHREPETITION": 9,"CYCLICIFACTIVEANDSPONTANWD": 10,"CYCLICIFACTIVEFAST": 11,"CYCLICWITHREPEATONDEMAND": 12,"NOSIGSENDTYPE": 13,"NOMSGSENDTYPE": 14,"NOTUSED": 15,"RESERVED": 16}
GenMsgSendType_List  = ["Cyclic",   "IfActive",   "cyclicIfActive",   "spontanWithDelay",   "cyclicAndSpontanX",   "cyclicAndSpontanWithDelay",   "spontanWithRepetition",   "cyclicIfActiveAndSpontanWD",   "cyclicIfActiveFast",   "cyclicWithRepeatOnDemand",   "NoSigSendType",    "NoMsgSendType",    "NotUsed",    "reserved"]
GenMsgSendType_Order = {"CYCLIC": 0,"IFACTIVE": 1,"CYCLICIFACTIVE": 2,"SPONTANWITHDELAY": 3,"CYCLICANDSPONTANX": 4,"CYCLICANDSPONTANWITHDELAY": 5,"SPONTANWITHREPETITION": 6,"CYCLICIFACTIVEANDSPONTANWD": 7,"CYCLICIFACTIVEFAST": 8,"CYCLICWITHREPEATONDEMAND": 9,"NOSIGSENDTYPE": 10,"NOMSGSENDTYPE": 11,"NOTUSED": 12,"RESERVED": 13}

SigType_List         = ["Default",    "Range",    "RangeSigned",    "ASCII",    "Discrete",    "Control",    "ReferencePGN",    "DTC",    "StringDelimiter",    "StringLength",    "StringLengthControl",     "MessageCounter",     "MessageChecksum"]
SigType_Order        = {"DEFAULT": 0, "RANGE": 1, "RANGESIGNED": 2, "ASCII": 3, "DISCRETE": 4, "CONTROL": 5, "REFERENCEPGN": 6, "DTC": 7, "STRINGDELIMITER": 8, "STRINGLENGTH": 9, "STRINGLENGTHCONTROL": 10, "MESSAGECOUNTER": 11, "MESSAGECHECKSUM": 12}

GenSigSendType_List  = ["Cyclic",    "OnChange",    "OnWrite",    "IfActive",    "OnChangeWithRepetition",    "OnWriteWithRepetition",    "IfActiveWithRepetition",    "NoSigSendType"]
GenSigSendType_Order = {"CYCLIC": 0, "ONCHANGE": 1, "ONWRITE": 2, "IFACTIVE": 3, "ONCHANGEWITHREPETITION": 4, "ONWRITEWITHREPETITION": 5, "IFACTIVEWITHREPETITION": 6, "NOSIGSENDTYPE": 7}

VFrameFormat_List    = ["StandardCAN",    "ExtendedCAN",    "reserved",    "J1939PG"]
VFrameFormat_Order   = {"STANDARDCAN": 0, "EXTENDEDCAN": 1, "RESERVED": 2, "J1939PG": 3}

EnumerationListTypes    = {'DiagRequest': Boolean_List,  'DiagResponse': Boolean_List,  'DiagState': Boolean_List,  'GenMsgSendType': GenMsgSendType_List,  'GenMsgILSupport': Boolean_List,  'NmNode'      : NmNode_List,
                           'NmMessage':   Boolean_List,  'ILUsed':       Boolean_List,  'SigType':   SigType_List,  'GenSigSendType': GenSigSendType_List,  'GenSigILSupport': Boolean_List,  'VFrameFormat': VFrameFormat_List}
EnumerationOrderTypes   = {'DiagRequest': Boolean_Order, 'DiagResponse': Boolean_Order, 'DiagState': Boolean_Order, 'GenMsgSendType': GenMsgSendType_Order, 'GenMsgILSupport': Boolean_Order, 'NmNode'      : NmNode_Order,
                           'NmMessage':   Boolean_Order, 'ILUsed':       Boolean_Order, 'SigType':   SigType_Order, 'GenSigSendType': GenSigSendType_Order, 'GenSigILSupport': Boolean_Order, 'VFrameFormat': VFrameFormat_Order}

# pre-defined attribution definitions
### CANdb++ complained that the attributes marked as mode=False had a "bad mode" or syntax error, thus not put into output files.
#  object type    name                   value type     min  max         default               mode    value range          values current file
ATTR_DEFS_INIT = [
    ["Message", "DiagRequest",           "Enumeration", "",  "",         "No",                 True,   Boolean_List],        # values for the 
    ["Message", "DiagResponse",          "Enumeration", "",  "",         "No",                 True,   Boolean_List],        # current file
    ["Message", "DiagState",             "Enumeration", "",  "",         "No",                 True,   Boolean_List],        # are not initialized
    ["Message", "GenMsgCycleTime",       "Integer",     0,   0,          0,                    True,   []],                  # loading each file
    ["Message", "GenMsgSendType",        "Enumeration", "",  "",         "NoMsgSendType",      True,   GenMsgSendType_List], # but are set when 
    ["Message", "GenMsgCycleTimeActive", "Integer",     0,   0,          0,                    True,   []],
    ["Message", "GenMsgCycleTimeFast",   "Integer",     0,   0,          0,                    True,   []],
    ["Message", "GenMsgDelayTime",       "Integer",     0,   0,          0,                    True,   []],
    ["Message", "GenMsgILSupport",       "Enumeration", "",  "",         "No",                 True,   Boolean_List],
    ["Message", "GenMsgNrOfRepetition",  "Integer",     0,   0,          0,                    True,   []],
    ["Message", "GenMsgStartDelayTime",  "Integer",     0,   65535,      0,                    True,   []],
    ["Message", "NmMessage",             "Enumeration", "",  "",         "No",                 True,   Boolean_List],
    ["Network", "BusType",               "String",      "",  "",         "CAN",                True,   []],
    ["Network", "Manufacturer",          "String",      "",  "",         "",                   True,   []],  ### ???
    ["Network", "NmBaseAddress",         "Hex",         0,  2047,        1024,                 True,   []],  ### Hex values are stored as integer, shown as int in dbc file
    ["Network", "NmMessageCount",        "Integer",     0,   255,        128,                  True,   []],
    ["Network", "NmType",                "String",      "",  "",         "",                   True,   []],
    ["Network", "DBName",                "String",      "",  "",         "CAN",                True,   []],
    ["Network", "Baudrate",              "Integer",     0,   1000000,    500000,               True,   []],
    ["Network", "DatabaseVersion",       "String",      "",  "",         "7.5",                False,  []],   ### CANdb++ doesn't seem to like this attribute
    ["Network", "ProtocolType",          "String",      "",  "",         "J1939",              True,   []],
    ["Network", "SAE_J1939_75_SpecVersion", "String",   "",  "",         "",                   True,   []],
    ["Network", "SAE_J1939_21_SpecVersion", "String",   "",  "",         "",                   True,   []],
    ["Network", "SAE_J1939_73_SpecVersion", "String",   "",  "",         "",                   True,   []],
    ["Network", "SAE_J1939_71_SpecVersion", "String",   "",  "",         "",                   True,   []],
    ["Node",    "DiagStationAddress",    "Hex",         0x0, 0xFF,       0x0,                  True,   []],
    ["Node",    "ILUsed",                "Enumeration", "",  "",         "No",                 True,   Boolean_List],
    ["Node",    "NmCAN",                 "Integer",     0,   2,          0,                    True,   []],
    ["Node",    "NmNode",                "Enumeration", "",  "",         "Not",                True,   NmNode_List],
    ["Node",    "NmStationAddress",      "Hex",         0x0, 0xFF,       0x0,                  True,   []],
    ["Node",    "NmSystemInstance",      "Integer",     0,   65535,      0,                    True,   []],
    ["Node",    "NmJ1939SystemInstance", "Integer",     0,   65535,      0,                    True,   []],
    ["Node",    "NodeLayerModules",      "String",      "",  "",         "CANoeILNVector.dll", True,   []],
    ["NodeMappedRXSignal", "GenSigTimeoutTime", "Integer", 0, 65535,     0,                    True,   []],
    ["Signal",  "SPN",                   "Integer",     0,   524287,     0,                    True,   []],
    ["Signal",  "GenSigStartValue",      "Integer",     0,   0,          0,                    True,   []],
    ["Signal",  "SigType",               "Enumeration", "",  "",         "",                   True,   SigType_List],
    ["Signal",  "GenSigInactiveValue",   "Integer",     0,   0,          0,                    True,   []],
    ["Signal",  "GenSigSendType",        "Enumeration", "",  "",         "",                   True,   GenSigSendType_List],
    ["Signal",  "GenSigTimeoutValue",    "Integer",     0,   1000000000, 0,                    True,   []],
    ["Signal",  "GenSigILSupport",       "Enumeration", "",  "",         "No",                 True,   Boolean_List],
    ["Signal",  "GenSigCycleTime",       "Integer",     0,   0,          0,                    True,   []],
    ["Signal",  "GenSigCycleTimeActive", "Integer",     0,   0,          0,                    True,   []],
    ["Signal",  "DDI",                   "Integer",     0,   65535,      0,                    True,   []],
    ["Signal",  "SAEDocument",           "String",      "",  "",         "J1939",              True,   []],
    ["Message", "VFrameFormat",          "Enumeration", 0,   8,          "ExtendedCAN",        True,   VFrameFormat_List],
    ["Signal",  "SystemSignalLongSymbol","String",      "",  "",         "STRING",             True,   []],
]

# matrix template dict
# Matrix parser use this dict to parse can network information from excel. The
# 'key's are used in this module and the 'ValueTable' are used to match excel
# column header. 'ValueTable' may include multi string.
MATRIX_TEMPLATE_MAP = {
    "msg_name_col":         ["MsgName"],
    "msg_type_col":         ["MsgType"],
    "msg_id_col":           ["MsgID"],
    "msg_send_type_col":    ["MsgSendType"],
    "msg_cycle_col":        ["MsgCycleTime"],
    "msg_len_col":          ["MsgLength"],
    "sig_name_col":         ["SignalName"],
    "sig_mltplx_col":       ["MultiplexerIndicator"],
    "sig_comment_col":      ["Description"],
    "sig_byte_order_col":   ["ByteOrder"],
    "sig_start_bit_col":    ["StartBit"],
    "sig_len_col":          ["BitLength"],
    "sig_value_type_col":   ["DataType"],
    "sig_factor_col":       ["Resolution"],
    "sig_offset_col":       ["Offset"],
    "sig_min_phys_col":     ["SignalMin.Value(phys)"],
    "sig_max_phys_col":     ["SignalMax.Value(phys)"],
    "sig_init_val_col":     ["InitialValue(Hex)"],
    "sig_unit_col":         ["Unit"],
    "sig_val_col":          ["SignalValueDescription"],
}

# excel workbook sheets with name in this list are ignored
MATRIX_SHEET_IGNORE = ["Cover", "History", "Legend", "ECU Version", ]

### changed from 8 to 16 based on an example file.
NODE_NAME_MAX = 16

def whoami():
    thisscript = os.path.basename(__file__)
    thisfunc = sys._getframe(1).f_code.co_name
    return (thisscript + ' ' + thisfunc + "():")

class MatrixTemplate(object):
    def __init__(self):
        self.msg_name_col = 0
        self.msg_type_col = 0
        self.msg_id_col = 0
        self.msg_send_type_col = 0
        self.msg_cycle_col = 0
        self.msg_len_col = 0
        self.sig_name_col = 0
        self.sig_mltplx_col = 0
        self.sig_comment_col = 0
        self.sig_byte_order_col = 0
        self.sig_start_bit_col = 0
        self.sig_len_col = 0
        self.sig_value_type_col = 0
        self.sig_factor_col = 0
        self.sig_offset_col = 0
        self.sig_min_phys_col = 0
        self.sig_max_phys_col = 0
        self.sig_init_val_col = 0
        self.sig_unit_col = 0
        self.sig_val_col = 0
        self.nodes = {}

        self.start_row = 0  # start row number of valid data

    def members(self):
        return sorted(vars(self).items(), key=lambda item:item[1] if type(item[1]) == int else None)

    def __str__(self):
        s = []
        for key, var in self.members():
            if type(var) == int:
                s.append("  %s : %d (%s)" % (key, var, get_xls_col(var)))
        return '\n'.join(s)


def get_xls_col(val):
    """
    get_xls_col(col) --> string
    
    Convert int column number to excel column symbol like A, B, AB .etc
    """
    if type(val) == type(0):
        if val <= 25:
            s = chr(val+0x41)
        elif val <100:
            s = chr(val/26-1+0x41)+chr(val%26+0x41)
        else:
            raise ValueError(whoami(), "column number too large: ", str(val))
        return s
    else:
        raise TypeError(whoami(), "column number only support int: ", str(val))


def get_list_item(list_object):
    """
    get_list_item(list_object) --> object(string)

    Show object list to terminal and get selection object from ternimal.
    """
    list_index = 0
    for item in list_object:
        print( "   %2d %s" %(list_index, item))
        list_index += 1
    while True:
        user_input = input()
        if user_input.isdigit():
            select_index = int(user_input)
            if select_index < len(list_object):
                return list_object[select_index]
            else:
                print (whoami(), "input over range")
        else:
            print (whoami(), "input invalid")


def parse_sheetname(workbook):
    """
    Get sheet name of can matrix in the xls workbook. 
    Only informations in this sheet are used.

    """
    sheets = []
    for sheetname in workbook.sheet_names():
        if sheetname == "Matrix":
            return sheetname
        if sheetname not in MATRIX_SHEET_IGNORE:
            sheets.append(sheetname)
    if len(sheets)==1:
        return sheets[0]
    elif len(sheets)>=2:
        print ("Select one sheet below:")
        # print ("  ","  ".join(sheets)
        # return raw_input()
        return get_list_item(sheets)
    else:
        print ("Select one sheet blow:")
        # print ("  ","  ".join(workbook.sheet_names())
        # return raw_input()
        return get_list_item(workbook.sheet_names())


def parse_template(sheet):
    """
    parse_template(sheet) -> MatrixTemplate
    
    Parse column headers of xls sheet and the result of column numbers is 
    returned as MatrixTemplate object
    """
    # find table header row
    header_row_num = 0xFFFF
    for row_num in range(0, sheet.nrows):
        if sheet.row_values(row_num)[0].find("MsgName") != -1:
            #print ("table header row number: %d" % row_num
            header_row_num = row_num
    if header_row_num == 0xFFFF:
        raise ValueError(whoami() +" Can't find \"Msg Name\" in this sheet")
    # get header info
    template = MatrixTemplate()
    for col_num in range(0, sheet.ncols):
        value = sheet.row_values(header_row_num)[col_num]
        if value is not None:
            value = value.replace(" ","")
            for col_name in MATRIX_TEMPLATE_MAP.keys():
                for col_header in MATRIX_TEMPLATE_MAP[col_name]:
                    if col_header in value and getattr(template,col_name)==0:
                        setattr(template, col_name, col_num)
                        break
    template.start_row = header_row_num + 1
    # get ECU nodes
    node_start_col = template.sig_val_col
    try: 
        for col in range(node_start_col, sheet.ncols):
            value = sheet.row_values(header_row_num)[col]
            if value is not None and value != '':
                value = value.replace(" ","")
                if len(value) <= NODE_NAME_MAX:
                    if value in template.nodes:
                        raise ValueError           ### added to exit if duplicate nodenames.
                    template.nodes[value] = col
    except ValueError:
        exit (whoami() + " ValueError: Duplicate value in Node Colums: {}".format(value))
    # print ("detected nodes: ", template.nodes
    return template


def parse_sig_vals(val_str):
    """
    parse_sig_vals(val_str) -> {valuetable}

    Get signal key:value pairs from string. Returns None if failed.
    """
    sigvals = {}
    if val_str is not None and val_str != '':
        token = re.split('[\;\:\n]+', val_str.strip())
        #print(token[1], len(token))
        if len(token) >= 2:
            if len(token) % 2 == 0:
                for i in range(0, len(token), 2):
                    try:
                        sigval = getint(token[i])
                        desc = token[i + 1].replace('\"','').strip()
                        #vals[desc] = val  # descriptions are not unique, e.g. many values being assigned as "Reserved"
                        sigvals[sigval] = desc
                        ####print(val, '', end='')
                    except ValueError:
                        print (whoami(), "waring: ignored signal value definition: " ,token[i], token[i+1])
                        raise
                ###print()
                return sigvals
            else:
                # print ("waring: ignored signal value description: ", val_str
                raise ValueError()
        else:
            raise ValueError(val_str)
    else:
        return None


def getint(strng, default=None):
    """
    getint(strng) -> int
    
    Convert string to int number. If default is given, default value is returned 
    while str is None.
    """
    if strng == '':
        if default==None:
            raise ValueError(whoami() + " None type object is unexpected")
        else:
            return default
    else:
        try:
            val = int(strng)
            return val
        except (ValueError, TypeError):
            try:
                strng = strng.replace('0x','')
                val = int(strng, 16)
                return val
            except:
                raise

#def motorola_msb_2_motorola_backward(start_bit, sig_size, frame_size):
#    msb_bytes            = start_bit//8
#    msb_byte_bit         = start_bit%8
#    byte_remain_bits     = msb_byte_bit + 1
#    remain_bits          = sig_size
#    backward_start_bit   = (frame_size *8 -1) - (8 * msb_bytes) - (7 - msb_byte_bit) 
#    while remain_bits > 0:
#        if remain_bits > byte_remain_bits:
#            remain_bits         -= byte_remain_bits
#            backward_start_bit  -= byte_remain_bits
#            byte_remain_bits    = 8
#        else:
#            backward_start_bit = backward_start_bit - remain_bits + 1 
#            remain_bits = 0
#    return backward_start_bit

class CanNetwork(object):
    def __init__(self, init=True):
        self.nodeobjects = []  ### list of node objects, each node contains comment and attributes.
        self.messages = []
        self.name = '' ###'CAN'
        self.val_tables = []        ### list of dictionaries - valtablename : { number: text, ...}
        self.version = ''
        self.new_symbols = NEW_SYMBOLS
        self.attr_defs = []
        self.attrs = {}              ###  list of attrs[namestring] = valuestring
        self.envvars = []
        self._filename = ''
        self.sg_mul_val_items = []   # "SG_MUL_VAL_" items are not parsed, but simply pulled in and dumped
        
        if init:
            self._init_attr_defs()
            for attr in self.attr_defs:
                if attr.name == "DBName":
                    self.attrs["DBName"] = attr.default
                    break

    def _init_attr_defs(self):
        for attr_def in ATTR_DEFS_INIT:
            self.attr_defs.append(CanAttribution(attr_def[1], attr_def[0], attr_def[2], attr_def[3], attr_def[4],
                                                 attr_def[5], attr_def[6], attr_def[7]))
            
    def __str__(self):
        # ! version
        lines = ['VERSION ' + r'""']
        lines.append('\n\n\n')

        # ! new_symbols
        lines.append('NS_ :\n')
        for symbol in self.new_symbols:
            lines.append('    ' + symbol + '\n')
        lines.append('\n')

        # ! bit_timming
        lines.append("BS_:\n")
        lines.append('\n')

        # ! nodes
        line = ['BU_:']
        if len(self.nodeobjects) > 0:
            for node in self.nodeobjects:
                line.append(node.name)
        lines.append(' '.join(line) + '\n')
        lines.append('\n')
        
        # ! Value tables, if any
        for valtbl in self.val_tables: 
            for valtablename in valtbl:
                line = ["VAL_TABLE_"]
                line.append(valtablename)
                myvalstring = ''
                for val_key in valtbl[valtablename]:
                    myvalstring += " " + str(val_key) + ' \"' + valtbl[valtablename][val_key] + '\"'
                line.append(myvalstring)
                lines.append(' '.join(line) + ';\n')
        lines.append('\n')

        # ! messages
        for msg in self.messages:
            lines.append(str(msg) + '\n\n')
        #lines.append('\n')

        # ! message transmitters
        for msg in self.messages: 
            if len(msg.transmitters) > 0:
                line = ["BO_TX_BU_"]
                line.append(str(msg.msg_id))
                line.append(':')
                line.append(msg.get_transmitters())
                lines.append(' '.join(line) + ';\n')
        lines.append('\n')
        
        # ! environment variables
        for envvar in self.envvars:
            lines.append(str(envvar) + '\n')
        lines.append('\n')

        # ! comments
        lines.append('''CM_ "Content processed by CANdb.py as forked by Soothsmith"''' + ";" + "\n")
        for node in self.nodeobjects: 
            if node.comment != '':
                line = ['CM_', 'BU_', node.name, '\"' + node.comment + '\";']
                lines.append(' '.join(line) + '\n')
        for msg in self.messages:
            comment = msg.comment
            if comment != "":
                line = ['CM_', 'BO_', str(msg.msg_id), '\"' + comment + '\";']
                lines.append(' '.join(line) + '\n')
            for sig in msg.signals:
                comment = sig.comment
                if comment != "":
                    line = ['CM_', 'SG_', str(msg.msg_id), sig.name, '\"' + comment + '\";']
                    lines.append(' '.join(line) + '\n')
        for envvar in self.envvars:
            comment = envvar.comment
            if comment != '':
                line = ['CM_', 'EV_', str(envvar.env_var_name), '\"' + comment + '\";']
                lines.append(' '.join(line) + '\n')

        # ! attribution defines
        for attr_def in self.attr_defs:
            if attr_def.mode == True:
                if attr_def.object_type == "NodeMappedRXSignal":
                    line = ["BA_DEF_REL_ BU_SG_REL_"]
                else:
                    line = ["BA_DEF_"]
                obj_type = attr_def.object_type
                if (obj_type == "Node"):
                    line.append("BU_")
                elif (obj_type == "Message"):
                    line.append("BO_")
                elif (obj_type == "Signal"):
                    line.append("SG_")
                ##elif (obj_type == "Network")
                line.append(" \"" + attr_def.name + "\"")
                val_type = attr_def.value_type
                if (val_type.upper() == "ENUMERATION"):
                    line.append("ENUM")
                    val_range = []
                    for val in attr_def.values:                  ### This list may be merged from multiple files, so values may not be the same in output files. 
                        val_range.append("\"" + val + "\"")
                    line.append(",".join(val_range) + ";")
                elif (val_type.upper() == "STRING"):
                    line.append("STRING" + " " + ";")
                elif (val_type.upper() == "HEX"):
                    line.append("HEX")     ### by definition, the HEX min/max are represented as integers
                    line.append(str(attr_def.min))
                    line.append(str(attr_def.max) + ";")
                elif (val_type.upper() == "INTEGER"):
                    line.append("INT")
                    line.append(str(attr_def.min))
                    line.append(str(attr_def.max) + ";")
                elif (val_type.upper() == "FLOAT"):
                    line.append("FLOAT")
                    strmin  = str(int(attr_def.min)) if (attr_def.min == float(int(attr_def.min))) else str(attr_def.min)
                    strmax  = str(int(attr_def.max)) if (attr_def.max == float(int(attr_def.max))) else str(attr_def.max)
                    line.append(strmin)
                    line.append(strmax + ';')
                lines.append(" ".join(line) + '\n')

        # ! attribution default values
        for attr_def in self.attr_defs:
            if attr_def.mode == True:
                if attr_def.object_type == "NodeMappedRXSignal":
                    line = ["BA_DEF_DEF_REL_"]
                else: 
                    line = ["BA_DEF_DEF_"]
                line.append(" \"" + attr_def.name + "\"")
                val_type = attr_def.value_type.upper()
                if (val_type == "ENUMERATION"):
                    line.append("\"" + attr_def.default + "\"" + ";")
                elif (val_type == "STRING"):
                    line.append("\"" + attr_def.default + "\"" + ";")
                elif (val_type == "HEX"):
                    line.append(str(attr_def.default) + ";")
                elif (val_type == "INTEGER"):
                    line.append(str(attr_def.default) + ";")
                elif (val_type == "FLOAT"):
                    default   = str(int(attr_def.default)) if (attr_def.default == float(int(attr_def.default))) else str(attr_def.default)
                    line.append(default + ";")
                else:
                    line.append('\"' + str(self.attrs_def.default) + '\";')
                lines.append(" ".join(line) + '\n')

        # ! attribution values of network
        for atrdefs in self.attr_defs:
            if atrdefs.object_type == "Network" and atrdefs.mode == True:
                for atr in self.attrs:
                    if atr == atrdefs.name:
                        val_type = attr_def.value_type.upper()
                        line = ["BA_"]
                        line.append('\"' + atr + '\"')
                        if val_type == "INTEGER":
                            attrvalue = self.attrs[atr]
                            ###strattrval  = str(int(attrvalue)) if (attrvalue == float(int(attrvalue))) else str(attrvalue)
                            line.append(str(attrvalue) + ';')
                        elif val_type == "STRING":
                            line.append('\"' + self.attrs[atr] + '\";')
                        elif val_type == "HEX":  ### HEX values are stored and represented as integers.
                            line.append(str('0x' + self.attrs[atr] + ';'))
                        else:
                            line.append('\"' + self.attrs[atr] + '\";')
                        lines.append(" ".join(line) + '\n')
                        
        # ! node attributes "BA_ attr BU_ node..."
        for node in self.nodeobjects:
            lines.append(str(node))
        
        # ! message attribution values
        for msg in self.messages:
            for attr_def in self.attr_defs:
                if attr_def.name in msg.attrs:
                    if (msg.attrs[attr_def.name] != ''):
                        line = ["BA_"]
                        line.append("\"" + attr_def.name + "\"")
                        line.append("BO_")
                        line.append(str(msg.msg_id))
                        if attr_def.value_type.upper() == "ENUMERATION":
                            # write enum index instead of enum value
                            line.append(str(attr_def.values.index(str(msg.attrs[attr_def.name]))) + ";")
                        elif attr_def.value_type.upper() == "STRING":
                            line.append('\"' + msg.attrs[attr_def.name] + "\";")
                        else:
                            line.append(str(msg.attrs[attr_def.name]) + ";")
                        lines.append(" ".join(line) + '\n')
                        
        # ! signal attribution values
        for msg in self.messages:
            for sig in msg.signals:
                for attr_def in self.attr_defs:
                    if attr_def.name in sig.attrs:
                        if (sig.attrs[attr_def.name] != ''):                            
                            line = ["BA_"] 
                            line.append('\"' + attr_def.name + '\"')
                            line.append("SG_")
                            line.append(str(msg.msg_id))
                            line.append(sig.name)
                            if (attr_def.value_type.upper() == "ENUMERATION"):
                                line.append(str(attr_def.values.index(str(sig.attrs[attr_def.name]))) + ';')
                            elif (attr_def.value_type.upper() == "STRING"):
                                line.append('\"' + sig.attrs[attr_def.name] + '\";')
                            else:
                                attrvalue = sig.attrs[attr_def.name]
                                strattrval  = str(int(attrvalue)) if (attrvalue == float(int(attrvalue))) else str(attrvalue)
                                line.append(str(strattrval) + ';')
                            lines.append(' '.join(line) + '\n')
        
        # ! Signal Value tables
        for msg in self.messages:
            for sig in msg.signals:
                if sig.values is not None and len(sig.values) >= 1:
                    line = ['VAL_']
                    line.append(str(msg.msg_id))
                    line.append(sig.name)
                    for key in sig.values:
                        line.append(str(key))
                        line.append('"' + sig.values[key] + '"')
                    line.append(';')
                    lines.append(' '.join(line) + '\n')

        # ! Envronment variable value tables
        for envvar in self.envvars:
            if len(envvar.values) > 0: 
                line = ['VAL_'] 
                line.append(str(envvar.env_var_name))
                line.append(envvar.get_values_str())
                lines.append(' '.join(line) + ';\n')

        # ! Signal value types
        for msg in self.messages:
            for sig in msg.signals:
                if sig.valtype is not None:
                    line = ['SIG_VALTYPE_']
                    line.append(str(msg.msg_id))
                    line.append(sig.name)
                    line.append(':')
                    line.append(str(sig.valtype))
                    lines.append(' '.join(line) + ';\n')
                    
        for msg in self.messages:
            for sig_group in msg.sig_groups:
                if sig_group.name != '':
                    line = ['SIG_GROUP_']
                    line.append(str(msg.msg_id))
                    line.append(str(sig_group))
                    lines.append(' '.join(line) + ';\n')

        for sg_mul_val_item in self.sg_mul_val_items:
            lines.append(sg_mul_val_item + '\n')
            
        return ''.join(lines)
    
#EnumerationListTypes    = {'DiagRequest': Boolean_List,  'DiagResponse': Boolean_List,  'DiagState': Boolean_List,  'GenMsgSendType': GenMsgSendType_List,  'GenMsgILSupport': Boolean_List, 
#                           'NmMessage':   Boolean_List,  'ILUsed':       Boolean_List,  'SigType':   SigType_List,  'GenSigSendType': GenSigSendType_List,  'GenSigILSupport': Boolean_List,  'VFrameFormat': VFrameFormat_List}
#EnumerationOrderTypes    = {'DiagRequest': Boolean_Order, 'DiagResponse': Boolean_Order, 'DiagState': Boolean_Order, 'GenMsgSendType': GenMsgSendType_Order, 'GenMsgILSupport': Boolean_Order, 
#                           'NmMessage':   Boolean_Order, 'ILUsed':       Boolean_Order, 'SigType':   SigType_Order, 'GenSigSendType': GenSigSendType_Order, 'GenSigILSupport': Boolean_Order, 'VFrameFormat': VFrameFormat_Order}
    def add_attr_def(self, name, object_type, value_type, minvalue, maxvalue, default, values=None):
        index = None
        for i in range (0, len(self.attr_defs)):
            if self.attr_defs[i].name.upper() == name.upper():
                index = i
                break
        if index is None:
            attr_def = CanAttribution(name, object_type, value_type, minvalue, maxvalue, default, True, values)
            print(whoami(), "Info: Adding attribute", name, "type", object_type, "from file.")
            self.attr_defs.append(attr_def)
        else:
            attr_name = self.attr_defs[index].name
            if object_type != '' and self.attr_defs[index].object_type != object_type:
                print(whoami(), "Info: Change object type for", name, "from", self.attr_defs[index].object_type, "to", object_type)
                self.attr_defs[index].object_type = object_type
            if value_type != '' and self.attr_defs[index].value_type  != value_type:
                print(whoami(), "Info: Change value type for", name, "from", self.attr_defs[index].value_type, "to", value_type)
                self.attr_defs[index].value_type  = value_type
            if minvalue != '' and self.attr_defs[index].min != minvalue:
                print(whoami(), "Info: Change minimum for", name, "from", self.attr_defs[index].min, "to", minvalue)
                self.attr_defs[index].min = minvalue
            if maxvalue != '' and self.attr_defs[index].max != maxvalue:
                print(whoami(), "Info: Change maximum for", name, "from", self.attr_defs[index].max, "to", maxvalue)
                self.attr_defs[index].max = maxvalue
            if default != '' and default != None:
                print(whoami(), "Info: Change default for", name, "from", self.attr_defs[index].default, " to ", default )
                self.attr_defs[index].default = default
            ### mode is left as it is
                                                                    # to interpret read enum values from subsequent records
            if self.attr_defs[index].value_type == "Enumeration":
                self.attr_defs[index].file_values = values          # if a file is being loaded, be sure to use file_values 
                if len(self.attr_defs[index].values) == 0:          # check enum list length for 0
                    self.attr_defs[index].values = values           # assign values as passed in because there are no previous values
                else:
                    #print(whoami(), attr_name, 'previous values =', self.attr_defs[index].values, 'new values =', values)
                    if attr_name in EnumerationListTypes.keys():   # if a known enumeration type
                        for value in values:
                            if value.upper() not in [val.upper() for val in EnumerationListTypes[attr_name]]: # compare to the appropriate value list
                                valueitemaslist = []
                                valueitemaslist.append(value)  # now it's a list of a single item
                                print(whoami(), " Info: adding value \'", value, "\' to attribute: ", name, " values ", str(self.attr_defs[index].values), sep='')
                                valueitemasset = set(valueitemaslist)  # now a set to 'union' them
                                self.attr_defs[index].values = list(set(self.attr_defs[index].values).union(valueitemasset))
                                ## must also add to the appropriate order types.
                                length = len(EnumerationOrderTypes[attr_name]) # current number of items in enum list/order 
                                EnumerationListTypes[attr_name].append(value)
                                EnumerationOrderTypes[attr_name][value.upper()] = length
                                #print(attr_name, value, " ", length, EnumerationListTypes[attr_name])
                    else:
                        # this is not one of the predefined enumeration lists, so just merge test best we can
                        self.attr_defs[index].values = list(set(self.attr_defs[index].values).union(set(values)))   # Merge the values
            else:
                if self.attr_defs[index].values != values:
                    print(whoami(), " info: changing default attribute value for ", name, " from ", self.attr_defs[index].values, " to ",values)
                    self.attr_defs[index].values = values
                
            if len(self.attr_defs[index].values) != 0:              # check enum list length for 0
                if attr_name in EnumerationListTypes.keys():  # If it is one of the attributes associated with enumeration types
                    # sort enumeration values based on the attribute type
                    self.attr_defs[index].values = sorted(self.attr_defs[index].values, key=lambda v: EnumerationOrderTypes[attr_name][v.upper()])

    def get_attr_def(self, name):
        ret = None
        for attr_def in self.attr_defs:
            if attr_def.name == name:
                ret = attr_def
                break
        return ret

    def append_message(self, canmessage):
        for msg in self.messages: 
            if msg.msg_id == canmessage.msg_id:
                msg.merge(canmessage)
                return msg      ### CanMessage object
        ### if the above didn't find and merge a message, then add the new one
        self.messages.append(canmessage)
        return canmessage
                
                
    def set_msg_attr(self, msg_id, attr_name, value):
        for msg in self.messages:
            if msg.msg_id == msg_id:
                msg.set_attr(attr_name, value)
                break

    def get_msg_attr(self, msg_id, attr_name):
        value = None
        for msg in self.messages:
            if msg.msg_id == msg_id:
                if msg.attrs.has_key(attr_name):
                    value = msg.attrs[attr_name]
                else:
                    attr_def = self.get_attr_def(attr_name)
                    if attr_def is not None:
                        value = attr_def.default
                        break
        return value

    def set_sig_group(self, msg_id, name, repetitions=1, signals=[]):
        for msg in self.messages:
            if msg.msg_id == msg_id:
                msg.add_sig_group(name, repetitions, signals)
                break

    def set_sig_attr(self, msg_id, sig_name, attr_name, attr_value):
        for msg in self.messages:
            if msg.msg_id == msg_id:
                for sig in msg.signals:
                    if sig.name == sig_name:
                        sig.set_attr(attr_name, attr_value)
                        break
                    
    def set_sig_valtype(self, msg_id, sig_name, valtype):
        for msg in self.messages:
            if msg.msg_id == msg_id:
                for sig in msg.signals:
                    if sig.name == sig_name:
                        ###print(msg_id, ": ", sig_name, ": ", valtype)
                        sig.valtype = int(valtype)
                        break

    def convert_attr_def_value(self, attr_def_name, value_str):
        '''
        Convert string to a value with AttributionDefinition's data type.

        Except:
            ValueError: attribuiton definition is not exit, or value type
                        of AttrDef is not support yet.
        '''
        attr_def_value_type = None
        for attr_def in self.attr_defs:
            if attr_def.name == attr_def_name:
                attr_def_value_type = attr_def.value_type.upper()
                attr_def_values     = [attr for attr in attr_def.file_values]     ### For incoming data, use the values as defined in the present file, not the merged values for output.
                if debug_enable: print(attr_def.name, attr_def_name, attr_def.values, attr_def_values)
                break
        if attr_def_value_type == 'INTEGER':
            value  = int(value_str)
        elif attr_def_value_type == 'FLOAT':
            value  = float(value_str)
        elif attr_def_value_type == 'STRING':
            value  = value_str.strip('\"')
        elif attr_def_value_type == 'ENUMERATION':
            if value_str.isdigit():
                value  = attr_def_values[int(value_str)]
                # e.g. SigType_List[SigType_Order[value.upper()]]
                try: 
                    value = EnumerationListTypes[attr_def.name][EnumerationOrderTypes[attr_def.name][value.upper()]] # convert read text to best camel-case
                except: 
                    pass
                    #print("tried to convert enum value for attribute", attr_def.name, "with value", value)

            else:
                value = value_str[1:-1]
        elif attr_def_value_type == 'HEX':   ### HEX values are stored and represented as integers
            value  = int(value_str)
        elif attr_def_value_type is None:
            raise ValueError(whoami() + "Undefined attribution definition: {}".format(attr_def_name))
        else:
            raise ValueError(whoami() + "Unkown attribution definition value type: {}".format(attr_def_value_type))
        return value

    def set_node_comment(self, nodename, comment):
        node_found = False
        for node in self.nodeobjects:
            if node.name == nodename:
                node.comment = comment
                node_found = True
                break
        if node_found == False: 
            nodeobject = Node(nodename, comment)
            self.nodeobjects.append(nodeobject)
            
    def set_node_attribute(self, nodename, attrname, attrvalue):
        node_found = False
        for node in self.nodeobjects:
            if node.name == nodename:
                node.attrs[attrname] = attrvalue
                node_found = True
                break
        if node_found == False: 
            nodeobject = Node(nodename, '', {attrname: attrvalue})
            self.nodeobjects.append(nodeobject)

    def set_sig_comment(self, msg_id, signame, comment):
        for msg in self.messages: 
            if msg.msg_id == msg_id:
                for signal in msg.signals:
                    if signal.name == signame: 
                        signal.set_comment(comment)
                        break
    
    def append_sig_comment(self, msg_id, signame, comment):
        for msg in self.messages: 
            if msg.msg_id == msg_id:
                for signal in msg.signals:
                    if signal.name == signame: 
                        signal.append_comment(comment)
                        break

    def set_msg_comment(self, msg_id, comment):
        for msg in self.messages: 
            if msg.msg_id == msg_id:
                msg.set_comment(comment)
                break

    def append_msg_comment(self, msg_id, comment):
        for msg in self.messages: 
            if msg.msg_id == msg_id:
                msg.append_comment(comment)
                break

    def set_msg_transmitters(self, msg_id, transmitters=[]):
        for msg in self.messages: 
            if msg.msg_id == msg_id:
                msg.transmitters = transmitters
                break

    def set_env_comment(self, ev_var_name, comment):
        for envvar in self.envvars: 
            if envvar.env_var_name == ev_var_name:
                envvar.set_comment(comment)
                break

    def append_env_comment(self, ev_var_name, comment):
        for envvar in self.envvars: 
            if envvar.env_var_name == ev_var_name:
                envvar.append_comment(comment)
                break

    def set_env_vals(self, ev_var_name = '', values = {}):
        for envvar in self.envvars:
            if envvar.env_var_name == ev_var_name:
                envvar.values = values
                break

    def sort(self, option='id'):
        messages = self.messages
        if option == 'id':
            # sort by msg_id, id is treated as string, NOT numbers, to keep the same with candb++
            messages.sort(key=lambda msg: str(msg.msg_id))
            for msg in messages:
                signals = msg.signals
                signals.sort(key=lambda sig: sig.start_bit)
        elif option == 'name':
            messages.sort(key=lambda msg: msg.name)
        else:
            raise ValueError(whoami() + "Invalid sort option \'{}\'".format(option))

    def load(self, path):
        print(whoami(), "Reading: ", path)
        dbcline = (l.rstrip('\n') for l in open(path, 'r'))  ### generator to read lines from the file helps with multiline comments
        
        for line in dbcline:
            line_rtrimmed = line.rstrip()                  ### will need line_rtrimmed when reading comments to preserve indent.
            line_trimmed = line_rtrimmed.lstrip()
            if line_trimmed == '':
                continue
            line_split = re.split('[\s\(\)\[\]\|\,\:\;]+', line_trimmed) ### needs better regex to account for optional multiplexor indicator: '', 'M', 'm num', num = 0, 1, 2, ...
            if len(line_split) >= 2:
                if line_split[0] == 'VERSION':
                    continue
                elif line_split[0] == 'NS_':
                    continue
                elif line_split[0] == 'BS_':
                    continue
                # Nodes
                elif line_split[0] == 'BU_':
                    new_node_list = []
                    #self.nodes = {nodename: '' for nodename in line_split[1:]}
                    new_node_list = line_split[1:]
                    existing_node_name_list = [node_object.name for node_object in self.nodeobjects]
                    for new_node_name in new_node_list: 
                        if new_node_name not in existing_node_name_list:
                            self.nodeobjects.append(Node(new_node_name))  # create list of Node objects

                # Message
                elif line_split[0] == 'BO_':
                    msg = CanMessage()
                    msg.msg_id = int(line_split[1])
                    msg.name   = line_split[2]
                    msg.dlc    = int(line_split[3])
                    msg.sender = line_split[4]
                    msg = self.append_message(msg)    # if merging files, need the pointer to the original message to append/merge signals
                # Signal
                elif line_split[0] == 'SG_':
                    sig = CanSignal()     ### create object from class
                    ###                      1=name 2=Multiplexer   3=start_bit 4=len 5=endian 6=sign   7=factor        8=offset              9=min             10=max         11=unit 12=receivers
                    ###                       (1)       (2)             (3)   (4)    (5)   (6)            (7)              (8)                 (9)               (10)            (11)     (12)
                    if debug_enable: print(line_trimmed)
                    match = re.match(r'SG_\s+(\w+)\s+([\w\d]*)\s*\:\s+(\d+)\|(\d+)\@(\d)([\+\-])\s+\(([eE\d\.\+\-]+)\,([eE\d\.\+\-]+)\)\s+\[([eE\d\.\-\+]+)\|([eE\d\.\-\+]+)]\s+\"(.*)\"\s+(.*)', line_trimmed)
                    if match:
                        sig.name          = match.group(1)
                        sig.mux_indicator = match.group(2) # saves '', 'M', or 'mx' where x is an int
                        sig.start_bit     = int(match.group(3))
                        sig.sig_len       = int(match.group(4))
                        sig.byte_order    = match.group(5)
                        sig.value_type    = match.group(6)
                        sig.factor        = match.group(7)
                        sig.offset        = match.group(8)
                        sig.min           = match.group(9)
                        sig.max           = match.group(10)
                        sig.unit          = match.group(11)
                        sig.receivers     = list(re.split('\s+', match.group(12))) # split receivers to list
                        sig.use_name      = msg.name.upper() == 'VECTOR__INDEPENDENT_SIG_MSG'
                        if debug_enable: print(str(sig))
                    msg.add_signal(sig)
                    
                # read in comments from Messages, Signals or Nodes
                elif line_split[0] == 'CM_':
                    comment_type = line_split[1]
                    match0 = re.match(r'CM_\s\".*\";', line_trimmed)
                    if match0:
                        next   # throw away the generic comment
                    new_comment = ''
                    match1 = re.match(r'CM_\s+(%s)\s+(\d*)\s*([\w\d_]*)\s*\"(.+?)\"\s*;$' %comment_type, line_trimmed) ### completed comment line ends with ";
                    if match1:
                        end = True
                        comment_type = match1.group(1)
                        something = match1.group(2)
                        message_id = int(something) if something != '' else None
                        comment_string = match1.group(4).replace("\r", "\n")
                        ##if debug_enable: print("has end:", str(message_id), match.group(3), comment_string)
                        # save comment to message/signal
                        if comment_type == "BU_":
                            nodename   = match1.group(3)
                            self.set_node_comment(nodename, comment_string)   ### nodes as list of objects
                        elif comment_type == "EV_":
                            envvar     = match1.group(3)
                            self.set_env_comment(envvar, comment_string)             
                        elif comment_type == "BO_":
                            self.set_msg_comment(message_id, comment_string)
                            #print("BO_ one-liner", message_id, ":", new_comment)
                        elif comment_type == "SG_":
                            signal_name = match1.group(3)
                            self.set_sig_comment(message_id, signal_name, comment_string)
                    else: 
                        #               1=comment_type 2 = msg_id, 3 = signal name, 4 = comment content
                        match2 = re.match(r'CM_\s+(%s)\s+(\d+)\s+([\w\d_]*)\s*\"(.+)$' %comment_type, line) ### start of comment line but no "; to finish it
                        if match2: 
                            end = False
                            comment_type = match2.group(1)
                            message_id = int(match2.group(2))
                            signal_name = match2.group(3)
                            comment_string = match2.group(4)
                            if debug_enable: print("no end:", comment_type, str(message_id), signal_name, comment_string)
                            new_comment += comment_string + '\n'
                                #self.set_sig_comment(message_id, signal_name, comment_string + '\n')
                            while(end==False):
                                line_rtrimmed = next(dbcline)                                                  ### Note - reading next line here
                                if line_rtrimmed == '' or line_rtrimmed == '\n':                               ### blank line in comment
                                    if debug_enable: print("cont:...", comment_type, str(message_id), signal_name, "null line")
                                    new_comment += '\n'
                                        #self.append_sig_comment(message_id, signal_name, '\n')
                                else: 
                                    match3 = re.match(r'^(?!CM_)(.*)\";$', line_rtrimmed)                      ### continued comment line does not start with CM_ and the "; is detected
                                    if match3:
                                        comment_string = match3.group(1)
                                        new_comment += comment_string
                                        end = True
                                        if debug_enable: print("final:  ", comment_type, str(message_id), signal_name, comment_string)
                                            #self.append_sig_comment(message_id, signal_name, comment_string)
                                    else: 
                                        match4 = re.match(r'^(?!CM_)(.+)(?!\";)$', line_rtrimmed)               ### continued comment line does not start with CM_ and no "; is detected
                                        if match4:
                                            comment_string = match4.group(1)
                                            new_comment += comment_string + '\n'
                                            if debug_enable: print("cont:  ", comment_type, str(message_id), signal_name, comment_string)
                                                #self.append_sig_comment(message_id, signal_name, comment_string + '\n')
                            if comment_type == "BU_":
                                nodename = match2.group(3)
                                self.set_node_comment(nodename, new_comment)   ### nodes as list of objects
                            elif comment_type == "EV_":
                                envvar   = match2.group(3)
                                self.set_env_comment(envvar, new_comment)                                        
                            elif comment_type == "BO_":
                                self.set_msg_comment(message_id, new_comment)
                                if debug_enable: print("BO_ ", message_id, ":", new_comment)
                            elif comment_type == "SG_":
                                self.set_sig_comment(message_id, signal_name, new_comment)
                            
                # Attribution Definition
                elif line_split[0] == 'BA_DEF_':
                    attr_def_object             = line_split[1]
                    if attr_def_object == 'BO_':
                        attr_def_object_type    = 'Message'
                        attr_def_name_offset    = 2
                    elif attr_def_object == 'SG_':
                        attr_def_object_type    = 'Signal'
                        attr_def_name_offset    = 2
                    elif attr_def_object == 'BU_':
                        attr_def_object_type    = 'Node'
                        attr_def_name_offset    = 2
                    else: # Network
                        attr_def_object_type    = 'Network'
                        attr_def_name_offset    = 1

                    attr_def_name               = line_split[attr_def_name_offset][1:-1] # remove quotation markes
                    attr_def_value_type         = line_split[attr_def_name_offset + 1]
                    attr_def_default            = ''

                    if attr_def_value_type == 'ENUM':
                        attr_def_value_type_str = 'Enumeration'
                        attr_def_value_min      = ''
                        attr_def_value_max      = ''
                        attr_def_values         = []
                        if debug_enable: print("ENUM LINE: ", end=''); print(line_split)
                        # attr_def_values         = map(lambda val:val[1:-1], line_split[attr_def_name_offset + 2:]) #remove "
                        for i in range(attr_def_name_offset + 2, len(line_split), 1):
                            enumvalue = line_split[i].replace('\"','').strip()
                            if enumvalue != '' and enumvalue != None:
                                attr_def_values.append(enumvalue)
                        if debug_enable: print("ENUM SAVED: ", end=''); print(attr_def_values)
                    elif attr_def_value_type == 'FLOAT':
                        attr_def_value_type_str = 'Float'
                        attr_def_value_min      = float(line_split[attr_def_name_offset + 2])
                        attr_def_value_max      = float(line_split[attr_def_name_offset + 3])
                        attr_def_values         = []
                    elif attr_def_value_type == 'INT':
                        attr_def_value_type_str = 'Integer'
                        attr_def_value_min      = int(float(line_split[attr_def_name_offset + 2]))
                        attr_def_value_max      = int(float(line_split[attr_def_name_offset + 3]))
                        attr_def_values         = []
                    elif attr_def_value_type == 'HEX':
                        attr_def_value_type_str = 'Hex'  ### HEX values are stored and represented as integers.
                        attr_def_value_min      = int(line_split[attr_def_name_offset + 2])
                        attr_def_value_max      = int(line_split[attr_def_name_offset + 3])
                        attr_def_values         = []
                    elif attr_def_value_type == 'STRING':
                        attr_def_value_type_str = 'String'
                        attr_def_value_min      = '' 
                        attr_def_value_max      = ''
                        attr_def_values         = []
                    else:
                        print(line_trimmed)
                        raise ValueError(whoami() + "Unkown attribution definition value type: {}".format(attr_def_value_type))

                    self.add_attr_def(attr_def_name, attr_def_object_type, attr_def_value_type_str, \
                                      attr_def_value_min, attr_def_value_max, attr_def_default, attr_def_values)
                # Attribution Definition Default Value
                elif line_split[0] == 'BA_DEF_DEF_' or line_split[0] == "BA_DEF_DEF_REL_":
                    attr_def_name               = line_split[1][1:-1] #remove "
                    attr_def_value_type_str     = None
                    attr_def_default            = self.convert_attr_def_value(attr_def_name, line_split[2])
                    attr_def                    = self.get_attr_def(attr_def_name)
                    if attr_def is not None and attr_def_default != None:
                        if attr_def_name in EnumerationListTypes.keys():
                            try: 
                                attr_def_default = EnumerationListTypes[attr_def.name][EnumerationOrderTypes[attr_def.name][attr_def_default.upper()]] # convert read text to best camel-case
                            except: 
                                print("BA_DEF_DEF_ ... tried to convert enum value for attribute", attr_def.name, "with value", attr_def_default)
                        if not (str(attr_def.default) == str(attr_def_default)):
                            if str(attr_def.default) != '':
                                print(whoami(), "Info: Update Attribute default for \'", attr_def_name, "\' from \'", attr_def.default, "\' to \'", attr_def_default, "\'", sep='')
                            attr_def.default    = attr_def_default
                    else: 
                        print(whoami(), "Info: Threw away default value", attr_def_default, " for", attr_def_name, "which is not defined.")
                    
                # Object Attribution definition
                elif line_split[0] == 'BA_':
                    attr_name           = line_split[1][1:-1] # remove "
                    attr_object         = line_split[2]
                    if attr_object == 'BO_':
                        attr_msg_id     = int(line_split[3])
                        attr_value      = self.convert_attr_def_value(attr_name, line_split[4])
                        if attr_value is not None:
                            self.set_msg_attr(attr_msg_id, attr_name, attr_value)
                        else:
                            raise ValueError(whoami() + "Msg \'{}\' attribuition \'{}\' value is {}".format(attr_msg_id, attr_name, line_split[4])) 
                    elif attr_object == 'SG_':
                        attr_msg_id     = int(line_split[3])
                        signal_name     = line_split[4]
                        attr_value      = self.convert_attr_def_value(attr_name, line_split[5])
                        if attr_value is not None: 
                            self.set_sig_attr(attr_msg_id, signal_name, attr_name, attr_value)
                    elif attr_object == 'BU_':
                        nodename = line_split[3].strip()
                        attr_value = line_split[4]
                        ##attr_value      = self.convert_attr_def_value(attr_name, line_split[4])  ### should handle multiple attribute types, only handling as text string right now. # TO DO
                        if debug_enable: print("Attribute: ", attr_name, ",Node:", nodename, ",attr_value:", attr_value)
                        self.set_node_attribute(nodename, attr_name, attr_value)
                    else: 
                        # network attribute: 
                        attr_name       = line_split[1][1:-1]
                        attr_value      = line_split[2].strip('\"')
                        self.attrs[attr_name] = attr_value
                        
                # Value Tables
                elif line_split[0] == 'VAL_' or line_split[0] == 'VAL_TABLE_': 
                    match20 = re.match(r'(VAL_|VAL_TABLE_)\s+(\d*)\s*(\w+)\s+(.+)$', line)
                    if match20:
                        valtabletype = match20.group(1)
                        val_msg_id   = match20.group(2).strip()   ### will be None for Env values, integer for signal values
                        val_msg_id   = int(val_msg_id) if re.match(r'\d+',val_msg_id) else None
                        text         = match20.group(3)   ### will be the Signal name or the Env variable name     
                        value_text   = match20.group(4)   ### unparsed value list - be sure to get them all before parsing
                        match22 = re.search(';$', value_text.rstrip())
                        end = False
                        if match22:
                            end = True
                            value_text = value_text.replace(r';$', r'')
                        #if valtabletype == 'VAL_TABLE_': print("End:", end, "vtype:", valtabletype, "id:", val_msg_id, 'name', text, "values:", value_text)
                        while end==False:
                            line = next(dbcline)    # get the next line to add to the attribute list
                            match21 = re.match(r'(.*)$', line)
                            if match21: 
                                value_text_ext = match21.group(1)
                                match23 = re.search(';$', value_text_ext.rstrip())
                                if match23: 
                                    end = True
                                    value_text_ext = value_text_ext.replace(r';$',r'')
                                value_text += value_text_ext
                                #if valtabletype == 'VAL_TABLE_':print(end, "summed:", value_text)
                        values = {}
                        value_list = re.split(r'\"', value_text)
                        if ';' in value_list: value_list.remove(';')
                        if ''  in value_list: value_list.remove('')
                        #if valtabletype == 'VAL_TABLE_': print(valtabletype, val_msg_id, "valuelist", value_list)
                        #print(value_list)
                        try:
                            values   = {int(value_list[i].strip()): value_list[i+1] for i in range(0, len(value_list)-1, 2)} 
                        except:
                            raise                            
                        if valtabletype == 'VAL_':
                            if val_msg_id == None:        ### Environment Variable value list
                                env_var_name = text
                                self.set_env_vals(env_var_name, values)
                            else:                         ### Message/Signal value list
                                val_sig_name = text
                                self.set_sig_attr(val_msg_id, val_sig_name, 'values', values)
                        elif valtabletype == 'VAL_TABLE_': 
                            # reserved for VAL_TABLE_ type
                            vt_name = text
                            namedvaltbl = {vt_name: values}
                            self.val_tables.append(namedvaltbl)
                        
                elif line_split[0] == 'EV_':   ### environment variables
                    ###                  1-name     2-type       3-minimum        4-maximum            5-units         6-initval         7-id   8-acc-type
                    match10 = re.match(r'EV_\s+(\w+)\:\s+(\d)\s+\[([eE\d\.\+\-]+)\|([eE\d\.\+\-]+)\]\s+\"(.*)\"\s+([eE\d\.\+\-]+)\s(\d+)\s(\w+\d)\s+(.+)\s*;', line_trimmed)
                    if match10:
                        env_var_name    = match10.group(1)
                        env_var_type    = int(match10.group(2))
                        env_var_min     = float(match10.group(3))
                        env_var_max     = float(match10.group(4))
                        env_var_units   = match10.group(5)
                        env_var_init    = float(match10.group(6))
                        env_var_id      = int(match10.group(7))
                        env_var_acctype = match10.group(8)
                        env_var_nodes   = match10.group(9).split(r'\s+')
                        if debug_enable: print(env_var_name, "...", env_var_type, ",", env_var_min, ",", env_var_max, ",", env_var_units, ",", env_var_init, ",", env_var_acctype, end=' ')
                        if debug_enable: print(' '.join(env_var_nodes), end='\n')
                        ###                  name='',      ev_id= 0,   env_var_type=0, units='',      minimum=0,   maximum=0,  initial_value=0, access_type=0,   access_nodes=[]):
                        envvar = EnvVariable(env_var_name, env_var_id, env_var_type,   env_var_units, env_var_min, env_var_max, env_var_init,   env_var_acctype, env_var_nodes)
                        self.envvars.append(envvar)
                        
                elif line_split[0] == 'BO_TX_BU_':
                    match11 = re.match(r'BO_TX_BU_\s+(\d+)\s*\:\s+(.+);', line_trimmed)
                    if match11:
                        msg_id = int(match11.group(1))
                        trans_msg_transmitters = match11.group(2).split(r'\,+')
                        self.set_msg_transmitters(msg_id, trans_msg_transmitters)
                        
                elif line_split[0] == 'SIG_VALTYPE_':
                    match12  = re.match(r'SIG_VALTYPE_\s+(\d+)\s+(\w+)\s+\:\s+([\w\d]+)\s*;', line_trimmed)
                    if match12:
                        msg_id = int(match12.group(1))
                        signame = match12.group(2)
                        sigvaltype = int(match12.group(3))   ### 0:signed or unsigned integer, 1: 32-bit IEEE-float, 2: 64-bit IEEE-double
                        self.set_sig_valtype(msg_id, signame, sigvaltype)
                        
                elif line_split[0] == 'SIG_GROUP_':
                    match13 = re.match(r'SIG_GROUP_\s(\d+)\s+(\w+)\s+(\d+)\s+\:\s+(.+)\s*;', line_trimmed)
                    if match13:
                        msg_id         = int(match13.group(1))
                        sig_group_name = match13.group(2)
                        repetitions    = int(match13.group(3))
                        signals_string = match13.group(4)
                        signals = re.split(r'\s+', signals_string)
                        self.set_sig_group(msg_id, sig_group_name, repetitions, signals)
                        
                elif line_split[0] == 'SG_MUL_VAL_':
                    self.sg_mul_val_items.append(line_trimmed)
                    
                else:
                    print("Unparsed: ", line_trimmed)

        dbcline.close()
                        
    def save(self, path=None):
        if (path == None):
            file = open(self._filename + ".dbc", "w")
            print("printing to:", self._filename + ".dbc")
        else:
            file = open(path, 'w')
        # file.write(unicode.encode(str(self), "utf-8"))
        file.write(str(self))

    def import_excel(self, path, sheetname=None, template=None):
        # Open file
        book = xlrd.open_workbook(path)
        # open sheet
        if sheetname is not None:
            #print ("use specified sheet: %s"%sheetname)
            sheet = book.sheet_by_name(sheetname)
        else:
            sheetname = parse_sheetname(book)
            #print ("select sheet: ", sheetname
            sheet = book.sheet_by_name(sheetname)

        # import template
        if template is not None:
            #print 'use specified template: ', template
            import_string = "import templates." + template + " as template"
            exec (import_string)
        else:
            print (whoami(), "parse template")
            template = parse_template(sheet)
            #if debug_enable:
                #print (template)
        # ! load network information
        filename = os.path.basename(path).split(".")
        fullpath = os.path.dirname(os.path.abspath(path)) + "\\"
        self._filename = fullpath + ".".join(filename[:-1])
        #self.name = ".".join(filename[:-1]).replace(" ", "_").replace('.', '_').replace('-', '_')  # use filename as default DBName
        self.name = "CAN"
        
        # ! load nodes information
        self.nodeobjects = {Node(nodekey) for nodekey in template.nodes.keys()}

        # ! load messages
        messages = self.messages
        nrows = sheet.nrows
        for row in range(template.start_row, nrows):
            row_values = sheet.row_values(row)
            msg_name = row_values[template.msg_name_col]
            if (msg_name != ''):
                # This row defines a message!
                message = CanMessage()
                signals = message.signals
                message.name = msg_name.replace(' ', '')
                message.msg_id = getint(row_values[template.msg_id_col])
                message.dlc = getint(row_values[template.msg_len_col])
                message_type = row_values[template.msg_type_col].upper().strip()
                if message_type == "J1939 PG (EXT. ID)":
                    message.set_attr("VFrameFormat", "J1939PG")
                elif message_type == "CAN STANDARD":
                    message.set_attr("VFrameFormat", "StandardCAN")
                elif message_type == "CAN EXTENDED":
                    message.set_attr("VFrameFormat", "ExtendedCAN")
                send_type = row_values[template.msg_send_type_col].upper().strip()
                if (send_type.upper() == "CYCLIC") or (send_type == "CE"):  # todo: CE is treated as cycle
                    try:
                        msg_cycle = getint(row_values[template.msg_cycle_col])
                    except ValueError:
                        print (whoami(), "warning: message %s\'s cycle time \"%s\" is invalid, auto set to \'0\'" % (message.name, row_values[template.msg_cycle_col]))
                        msg_cycle = 0
                    message.set_attr("GenMsgCycleTime", msg_cycle)
                    message.set_attr("GenMsgSendType", "Cyclic")
                elif (send_type.upper() == "NOMSGSENDTYPE"):
                    message.set_attr("GenMsgSendType", "NoMsgSendType")
                else:
                    message.set_attr("GenMsgSendType", "NoMsgSendType")
                message.comment = row_values[template.sig_comment_col].replace("\"", "\'").replace(";",",").replace("\r", '\n').replace("\n\n", "\n")
                # message sender
                message.sender = None
                for nodename in template.nodes:
                    nodecol = template.nodes[nodename]
                    sender = row_values[nodecol].strip().upper()
                    if sender == "S":
                        message.sender = nodename
                    elif sender == "R":
                        message.receivers.append(nodename)
                if message.sender is None:
                    message.sender = 'Vector__XXX'
                messages.append(message)
            else:
                sig_name = row_values[template.sig_name_col]
                if (sig_name != ''):
                    # This row defines a signal!
                    signal = CanSignal()
                    signal.set_compare_type(msg_name.upper() =='VECTOR__INDEPENDENT_SIG_MSG')  # repeat signal names allowed in this msg
                    signal.name = sig_name.replace(' ', '')
                    mux_value = row_values[template.sig_mltplx_col]
                    signal.mux_indicator = mux_value if (mux_value == 'M' or mux_value == '') else 'm' + str(int(mux_value))
                    ###print(signal.mux_indicator)
                    signal.start_bit = getint(row_values[template.sig_start_bit_col])
                    signal.sig_len = getint(row_values[template.sig_len_col])
                    
                    byte_order_type = row_values[template.sig_byte_order_col]
                    if (byte_order_type.upper() == "MOTOROLA LSB"):  # todo
                        signal.byte_order = '1'
                    elif (byte_order_type.upper() == "MOTOROLA MSB"):
                        signal.byte_order = '0'
                    else:
                        signal.byte_order = '0'
                        raise ValueError(whoami() + "Unknown signal byte order type: \"%s\""%byte_order_type) # todo: intel
                    
                    signal_value_type = row_values[template.sig_value_type_col].upper()
                    if (signal_value_type == "UNSIGNED"):
                        signal.value_type = '+'
                        signal.valtype    = None     # not float
                    elif (signal_value_type == "SIGNED"):
                        signal.value_type = '-'
                        signal.valtype    = None     # not float
                    elif (signal_value_type == "IEEE FLOAT"):
                        signal.value_type = '-'
                        signal.valtype    = 1     # 32-bit IEEE float
                    elif (signal_value_type == "IEEE DOUBLE"):
                        signal.value_type = '-'
                        signal.valtype    = 2     # 64-bit IEEE float
                        
                    signal.factor = row_values[template.sig_factor_col]
                    signal.offset = row_values[template.sig_offset_col]
                    signal.min = row_values[template.sig_min_phys_col]
                    signal.max = row_values[template.sig_max_phys_col]
                    signal.unit = row_values[template.sig_unit_col]
                    signal.attrs["GenSigStartValue"] = getint((row_values[template.sig_init_val_col]), 0)
                    try:
                        signal.values = parse_sig_vals(row_values[template.sig_val_col])
                    except ValueError:
                        if debug_enable:
                            print (whoami(), "warning: signal %s\'s value table is ignored" % signal.name)
                    signal.comment = row_values[template.sig_comment_col].replace("\"", "\'").replace(";",",").replace("\r", '\n').replace("\n\n", "\n") # todo
                    # get signal receivers
                    signal.receivers = []
                    for nodename in template.nodes:
                        nodecol = template.nodes[nodename]
                        receiver = row_values[nodecol].strip().upper()
                        if receiver == "R":
                            signal.receivers.append(nodename)
                        elif receiver == "S":
                            if message.sender == 'Vector__XXX':
                                message.sender = nodename
                                print (whoami(), "warning: message %s\'s sender is set to \"%s\" via signal" %(message.name, nodename))
                            else:
                                print (whoami(), "warning: message %s\'s sender is conflict to signal \"%s\"" %(message.name, nodename))
                        else:
                            pass
                    if len(signal.receivers) == 0:
                        if len(message.receivers) > 0:
                            signal.receivers = message.receivers
                        else:
                            signal.receivers.append('Vector__XXX')
                    signals.append(signal)
        self.sort();

class Node(object):
    '''
    There would be a list of nodes, each having a comment and attributes.
    -- A list of nodes is kept with the CanNetwork object
    '''
    def __init__(self, name='', comment='', attrs=None):
        self.name = name
        self.comment = comment
        if attrs == None: 
            self.attrs = {}
        else: 
            self.attrs = attrs
        if debug_enable: print("new node created:", name, "comment:", comment, "attrs:", str(attrs))
    
    def __str__(self):
        strnodeattributes = ''
        for attr in self.attrs:
            ### compare to attrtable for types.
            line = ["BA_"]
            line.append('\"' + attr + '\"')
            line.append("BU_")
            line.append(self.name)
            line.append(str(self.attrs[attr]))
            strnodeattribute = (' '.join(line) + ';\n')
            strnodeattributes = strnodeattributes + strnodeattribute
        return(strnodeattributes)
    
    def set_node_attribute(self, attrname, attrvalue):
        self.attrs[attrname] = attrvalue
        #print("new node attribute: ", self.name, "attr:", attrname, "attrvalue:", attrvalue)

class SigGroup(object):
    '''
    A signal group is "to define that the signals of 
    a group have to be updated in common."
    Denoted as SIG_GROUP_ in the dbc file, but this requires the 
    message ID specified, so only part of the string is done here.
    '''
    def __init__(self, name='', repetitions=1, signals=[]):
        self.name = name                   ### string
        self.repetitions = repetitions     ### probably left as 1, not sure what repetitions means
        self.signals = signals             ### a list of signal names belonging to a message
        
    def __str__(self):
        strsignals = ' '.join(self.signals)
        return(self.name + ' ' + str(self.repetitions) + ' : ' + strsignals)

class CanMessage(object):
    def __init__(self, name='', msg_id=0, dlc=8, sender='Vector__XXX'):
        ''' 
        name: message name
        id: message id (11bit or 29 bit)
        dlc: message data length
        sender: message send node
        '''
        self.name = name           ### string
        self.msg_id = msg_id       ### integer
        self.send_type = ''
        self.dlc = dlc
        self.sender = sender       ### string
        self.signals = []          ### list of CanSignal objects
        self.attrs = {}            ### dictionary
        self.receivers = []        ### list
        self.transmitters = []     ### a list for BO_TX_BU_
        self.comment = ''          ### comment comes in separately from the message definition
        self.sig_groups = []       ### a list of SigGroup objects

    def __str__(self):
        para = []
        line = ["BO_", str(self.msg_id), self.name + ":", str(self.dlc), self.sender]
        para.append(" ".join(line))
        for sig in self.signals:
            para.append(str(sig))
        return '\n '.join(para)
    
    def merge(self, canmessage):
        '''
        Messages are usually merged when a Message is read from a file, which is 
        before the signals are read. But if signals are defined, they should be merged in,
        though the signals may be in no particular order.
        '''
        print(whoami(), "info: Merge message: id=", self.msg_id, "Name:", self.name, end='')
        self.name                   = canmessage.name
        self.send_type              = canmessage.send_type
        self.dlc                    = canmessage.dlc
        if len(canmessage.sender) and canmessage.sender != 'Vector__XXX':  ### only overwrite if new value is a named ECU
            print(" sender", self.sender, "overwritten wtih", canmessage.sender, end='')
            self.sender             = canmessage.sender
        else:
            print()
        if len(canmessage.signals): ### if new message has signals to merge in
            if len(self.signals):  ### if stored message has signals
                for signal in self.signals:       ### walk through, find matches and merge
                    for newsig in canmessage.signals:
                        if signal == newsig:           ### operator overriden, selective fields compared.
                            signal.merge(newsig)
            else:  ## just assign new signals into existing message
                self.signals = canmessage.signals
        self.attrs.update(canmessage.attrs)                                                    ### merge dictionaries
        if len(canmessage.receivers):
            self.receivers         = list(set(self.receivers.extend(canmessage.receivers)))    ### merging list to have unique values
        if len(canmessage.transmitters):
            self.transmitters      = list(set(self.transmitters.extend(canmessage.transmitters))) ## merging list of strings
        if len(canmessage.sig_groups):
            self.sig_groups        = list(set(self.sig_groups.extend(canmessage.sig_groups)))  ### is a list of multilayer dictionaries...needs more work
        if len(self.comment) == 0:
            self.comment           = canmessage.comment.rstrip()                               ### comments probably needs more work.
        
    def add_signal(self, cansignal):
        for signal in self.signals:
            if signal == cansignal:    ### == has been redefined CanSignal                
                signal.merge(cansignal)
                return
        ### did not find a matching signal in this message, append new
        use_name = self.name.upper() == 'VECTOR__INDEPENDENT_SIG_MSG'
        cansignal.set_compare_type(use_name)
        self.signals.append(cansignal)
                
    def set_attr(self, name, value):
        self.attrs[name] = value
        
    def set_comment(self, comment):
        comment = comment.strip()
        existing_comment = self.comment
        x = existing_comment.find(comment) == None    ### existing comment is found, skip
        if not x:
            self.comment = comment
        else:
            pass
        ### self.comment = comment
            
    def append_comment(self, comment):
        self.comment = self.comment + comment

    def get_transmitters(self):
        return(','.join(self.transmitters))
    
    def str_sig_groups(self):
        lines = ''
        for sig_group in self.sig_group:
            line = ['SIG_GROUP_']
            line.append(str(self.msg_id))
            line.append(str(sig_group))
            lines = lines + ' '.join(line) + ';\n'
        return lines
    
    def add_sig_group(self, signal_group_name, repetitions, signals=[]):
            signal_group = SigGroup(signal_group_name, repetitions, signals)
            self.sig_groups.append(signal_group)

class CanSignal(object):
    def __init__(self, name='', start_bit=0, sig_len=1, init_val=0, use_name=False):
        self.name = name           ### string
        self.mux_indicator = ''    ### None, or 'M', or m0, m1, m2, m3, ...
        self.start_bit = start_bit ### integer, start bit
        self.sig_len = sig_len     ### integer, signal length
        self.valtype = None        ### 0:signed or unsigned integer, 1: 32-bit IEEE-float, 2: 64-bit IEEE-double
        self.byte_order = '0'      ### 1, 0
        self.value_type = '+'      ### '+' unsigned, '-' signed
        self.factor = 1            ### float, * scale factor
        self.offset = 0            ### float, + offset, float
        self.min = 0               ##3 float
        self.max = 1               ### float
        self.unit = ''             ### text units
        self.values = {}           ### dictionary: VAL_ table describes integer: value
        self.receivers = []        ### node names, list of strings (not the Node object)
        self.comment = ''          ### text, CM_ SG_ comment information
        self.attrs = {}            ### dictionary, BA_ SG_ attribute information
        self.use_name = use_name   ### use signal name in comparison for signals
        if init_val != 0:
            self.attrs["GenSigStartValue"] = init_val

    def __str__(self):
        line = ["SG_", self.name + ' '*bool(self.mux_indicator) + self.mux_indicator, ":",
                str(self.start_bit) + "|" + str(self.sig_len) + "@" + self.byte_order + self.value_type,
                "(" + str(self.factor) + "," + str(self.offset) + ")", "[" + str(self.min) + "|" + str(self.max) + "]",
                "\"" + str(self.unit) + "\"", ','.join(self.receivers)]
        return " ".join(line)
    
    def set_compare_type(self, use_name=False):
        self.use_name = use_name
        
    def __eq__(self, cansignal):
        '''
        Test whether CanSignal objects are equivalent. Several fields are tested, others ignored.
        '''
        return_val = False
        if (self.use_name):
            if (self.name          == cansignal.name
            and self.start_bit     == cansignal.start_bit         ### integer
            and self.sig_len       == cansignal.sig_len           ### number of bits
            and self.value_type    == cansignal.value_type        ### + unsigned, - signed
            and self.mux_indicator == cansignal.mux_indicator):   ### None, 'M', or m0, m1, m2, m3, ...
                return_val = True
        else: 
            if (self.start_bit     == cansignal.start_bit         ### integer
            and self.sig_len       == cansignal.sig_len           ### number of bits
            and self.value_type    == cansignal.value_type        ### + unsigned, - signed
            and self.mux_indicator == cansignal.mux_indicator):   ### None, 'M', or m0, m1, m2, m3, ...
                return_val = True
        return return_val
            
    def merge(self, cansignal):
        '''
        Signals have been compared and determined to be equivalent, 
        so merge the fields with some semblance of sanity.
        '''
        print(whoami(), "Info: Merge signal:", str(self), "with", str(cansignal))
        ## at the time the signal is read, the commonts nor attributes will have values, they come separately
        if len(self.name) < len(cansignal.name):   ### take the more verbose name
            self.name = cansignal.name
        self.mux_indicator         = cansignal.mux_indicator     ### same if __eq__ used
        self.start_bit             = cansignal.start_bit         ### integer
        self.sig_len               = cansignal.sig_len           ### number of bits
        if (self.valtype == None) and (cansignal.valtype != None):
            self.valtype           = cansignal.valtype           ### 0:signed or unsigned integer, 1: 32-bit IEEE-float, 2: 64-bit IEEE-double
        self.byte_order            = cansignal.byte_order        ### int, 1/0
        self.value_type            = cansignal.value_type        ### + unsigned, - signed
        self.min                   = cansignal.min               ### 
        self.max                   = cansignal.max               ### 
        self.unit                  = cansignal.unit              ### 
        if len(self.values) or len(cansignal.values):  ### merge dictionary values
            self.values.update(cansignal.values)
        if len(self.receivers) or len(cansignal.receivers):
            pass
            #self.receivers         = list(set(self.receivers.extend(cansignal.receivers)))  ### merging list to have unique values
        if len(self.comment) == 0:
            self.comment           = cansignal.comment.rstrip()                                   ### probably needs more work.            
        if len(self.attrs) or len(cansignal.attrs):  ### merge dictionary values
            self.values.update(cansignal.values)
            
    def set_attr(self, name, value):
        if name == 'values':
            self.values = value
        elif name in SignalAttributes_List: ###["SPN", "DDI", "SigType", "GenSigInactiveValue", "GenSigSendType", "GenSigStartValue", "GenSigTimeoutValue", "SAEDocument", "SystemSignalLongSymbol", "GenSigILSupport", "GenSigCycleTime", "GenSigCycleTimeActive"]:
            self.attrs[name] = value
        else:
            raise ValueError("Unsupport set attr of \'{}\'".format(name))

    def get_attr(self, name):
        if name == 'values':
            return self.values
        elif name in SignalAttributes_List:   ###["SPN", "DDI", "SigType", "GenSigInactiveValue", "GenSigSendType", "GenSigStartValue", "GenSigTimeoutValue", "SAEDocument", "SystemSignalLongSymbol", "GenSigILSupport", "GenSigCycleTime", "GenSigCycleTimeActive"]:
            return self.attrs[name]
        elif name == '':
            return self.value_type
        else:
            raise ValueError("Unsupport set attr of \'{}\'".format(name))

    def set_comment(self, comment):
        self.comment = comment
        
    def append_comment(self, comment):
        self.comment = self.comment + comment

class CanAttribution(object):
    def __init__(self, name, object_type, value_type, minvalue, maxvalue, default, mode=True, values=None):
        self.name = name
        self.object_type = object_type
        self.value_type = value_type
        self.min = minvalue
        self.max = maxvalue
        self.default = default
        self.mode = mode
        self.values = values
        self.file_values = values

    def __str__(self):
        line = ["name: {}".format(self.name)]
        line.append("object type: {}".format(self.object_type))
        line.append("value type: {}".format(self.value_type))
        line.append("minimum:   {}".format(self.min))
        line.append("maximum:   {}".format(self.max))
        line.append("default:   {}".format(self.default))
        line.append("values:")
        if self.value_type == "Enumeration" and self.values is not None:
            for i in range(0, len(self.values)):
                line.append("       {:d}:   {}".format(i, self.values[i]))
                line.append("values:")
        if self.value_type == "Enumeration" and self.values is not None:
            for i in range(0, len(self.values_from_file)):
                line.append("       {:d}:   {}".format(i, self.values[i]))
        return '\n'.join(line)

class EnvVariable(object):
    '''
    EnvVariable class is intended to hold an environment variable's attributes including a list of values.
    The spec isn't clear about what such a variable is, but they show up in files and would need to be read.
    '''
    access_type_text = {'unrestricted': 'DUMMY_NODE_VECTOR0', 'read': 'DUMMY_NODE_VECTOR1', 'write': 'DUMMY_NODE_VECTOR2', 'readWrite': 'DUMMY_NODE_VECTOR3'}
    
    str_access_type = lambda self, actp: EnvVariable.access_type_text[actp]
    
    def cvt_access_type_from_string(self, as_string):
        '''
        Convert the values like "DUMMY_NODE_VECTORX" to the understandable value for Env Variables.
        The intent is that if Env Variables are added to a sheet in the Excel workbook, the selection would be the readable items not "DUMMY..." 
        '''
        for ac_key, ac_val in self.access_type_text.items():
            if ac_val == as_string: 
                return ac_key

    def __init__(self, name='', ev_id= 0, env_var_type=0, units='', minimum=0, maximum=0, initial_value=0, access_type=0, access_nodes=[]):
        ''' 
        name: enviroment variable name
        ev_id: id -- which is apparently arbitrary.
        access_type is DUMMY_NODE_VECTORX but is stored as 'unrestricted'... 
        '''
        self.env_var_name = name            ### string
        self.ev_id = ev_id                  ### integer
        self.units = units                  ### string
        self.env_var_type = env_var_type ###if env_var_type in {0: 'Integer', 1: 'Float', 2: 'String'}.keys() else 0
        self.minimum = minimum              ### float
        self.maximum = maximum              ### float
        self.initial_value = initial_value  ### float
        self.access_type = self.cvt_access_type_from_string(access_type)
        self.access_nodes = access_nodes    ### list
        self.values = {}                    ### these values are part of the "VAL_" lines and not part of the "EV_" line
        self.comment = ''                   ### char string associated with the env var

    def __str__(self):
        strmin  = str(int(self.minimum)) if (self.minimum == float(int(self.minimum))) else str(self.minimum) ### drop decimal point if no meaningful content after it
        strmax  = str(int(self.maximum)) if (self.maximum == float(int(self.maximum))) else str(self.maximum)
        strinit = str(int(self.initial_value)) if (self.initial_value == float(int(self.initial_value))) else str(self.initial_value)
        line = ["EV_", self.env_var_name + ":", str(self.env_var_type), '[' + strmin + '|' + strmax + ']', '\"' + self.units + '\"', strinit, str(self.ev_id), self.str_access_type(self.access_type)]
        line.extend(self.access_nodes)
        return ' '.join(line) + ";\n"

    def set_value(self, name, value):
        self.values[value] = name    ### values unique, names may not be
        
    def get_values_str(self):
        '''
        return the formatted values for the EnvVariable's value table 
        '''
        myvalstring = ''
        for ev_val_key in self.values:
            myvalstring += " " + str(ev_val_key) + ' \"' + self.values[ev_val_key] + '\"'
        return myvalstring 
    
    def set_comment(self, comment):
        self.comment = comment.strip()
        
    def append_comment(self, comment):
        self.comment = self.comment + comment            

def parse_args():
    """
    Parse command line commands.
    """
    import argparse
    parse = argparse.ArgumentParser()
    subparser = parse.add_subparsers(title="subcommands")

    parse_gen = subparser.add_parser("gen", help="Generate dbc from excle file")
    parse_gen.add_argument("filename", help="The xls file to generate dbc")
    parse_gen.add_argument("-s","--sheetname",help="set sheet name of xls",default=None)
    parse_gen.add_argument("-t","--template",help="Choose a template",default=None)
    parse_gen.add_argument("-d","--debug",help="show debug info",action="store_true", dest="debug_switch", default=False)
    parse_gen.set_defaults(func=cmd_gen)

    parse_sort = subparser.add_parser("sort", help="Sort dbc messages and signals")
    parse_sort.add_argument("filename", help="Dbc filename")
    parse_sort.add_argument("-o","--output", help="Specify output file path", default=None)
    parse_sort.set_defaults(func=cmd_sort)

    parse_sort = subparser.add_parser("merge", help="Merge dbc messages and signals")
    parse_sort.add_argument("-f","--dbcfiles",   nargs="*", default=[], help="dbc filename list")
    parse_sort.add_argument("-o","--output", help="Specify output file path", default=None)
    parse_sort.set_defaults(func=cmd_merge)

    parse_cmp = subparser.add_parser("cmp", help="Compare difference bettween two dbc files - not yet implemented.")
    parse_cmp.add_argument("filename1", help="The base file to be compared with")
    parse_cmp.add_argument("filename2", help="The new file to be compared")
    parse_cmp.set_defaults(func=cmd_cmp)

    args = parse.parse_args()
    args.func(args)


def cmd_gen(args):
    global  debug_enable
    debug_enable = args.debug_switch
    try:
        can = CanNetwork()
        can.import_excel(args.filename, args.sheetname, args.template)
        can.save()
    except IOError as e:
        print (e)
    except xlrd.biffh.XLRDError as e:
        print (e)
        
        
def cmd_sort(args):
    can = CanNetwork()
    can.load(args.filename)
    can.sort()
    if args.output is None:
        can.save("sorted.dbc")
    else:
        can.save(args.output)
        
def cmd_merge(args):
    can = CanNetwork()
    for sourcefile in args.dbcfiles:
        can.load(sourcefile)
    if args.output is None:
        can.save("sorted.dbc")
    else: 
        can.save(args.output)


def cmd_cmp(args):
    print ("Compare function is comming soon!")


if __name__ == '__main__':
    parse_args()

