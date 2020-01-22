
"""\
This module provides CAN database(*dbc) and matrix(*xls) operation functions.


"""
import re
import xlrd
import sys, importlib
import traceback
import os
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


# pre-defined attribution definitions
#  object type    name                   value type     min  max         default               value range
ATTR_DEFS_INIT = [
    ["Message", "DiagRequest",           "Enumeration", "",  "",         "No",                 ["No",    "Yes"]],
    ["Message", "DiagResponse",          "Enumeration", "",  "",         "No",                 ["No",    "Yes"]],
    ["Message", "DiagState",             "Enumeration", "",  "",         "No",                 ["No",    "Yes"]],
    ["Message", "GenMsgSendType",        "Enumeration", "",  "",         "noMsgSendType",      ["cyclic", "reserved","cyclicIfActive","reserved","reserved","reserved","reserved","reserved","noMsgSendType"]],
    ["Message", "GenMsgCycleTime",       "Integer",     0,   0,          0,                    []],
    ["Message", "GenMsgCycleTimeActive", "Integer",     0,   0,          0,                    []],
    ["Message", "GenMsgCycleTimeFast",   "Integer",     0,   0,          0,                    []],
    ["Message", "GenMsgDelayTime",       "Integer",     0,   0,          0,                    []],
    ["Message", "GenMsgILSupport",       "Enumeration", "",  "",         "No",                 ["Yes",    "No"]],
    ["Message", "GenMsgNrOfRepetition",  "Integer",     0,   0,          0,                    []],
    ["Message", "GenMsgStartDelayTime",  "Integer",     0,   65535,      0,                    []],
    ["Message", "NmMessage",             "Enumeration", "",  "",         "No",                 ["No",    "Yes"]],
    ["Network", "BusType",               "String",      "",  "",         "CAN",                []],
    ["Network", "Manufacturer",          "String",      "",  "",         "",                   []],
    ["Network", "NmBaseAddress",         "Hex",         0x0, 0x7FF,      0x400,                []],
    ["Network", "NmMessageCount",        "Integer",     0,   255,        128,                  []],
    ["Network", "NmType",                "String",      "",  "",         "",                   []],
    ["Network", "DBName",                "String",      "",  "",         "CAN",                []],
    ["Network", "DatabaseVersion",       "Integer",     0,   0,          0,                    []],
    ["Network", "ProtocolType",          "String",      "",  "",         "J1939",              []],
    ["Network", "SAE_J1939_75_SpecVersion", "String",   "",  "",         "",                   []],
    ["Network", "SAE_J1939_21_SpecVersion", "String",   "",  "",         "",                   []], 
    ["Network", "SAE_J1939_73_SpecVersion", "String",   "",  "",         "",                   []],
    ["Network", "SAE_J1939_71_SpecVersion", "String",   "",  "",         "",                   []], 
    ["Node",    "DiagStationAddress",    "Hex",         0x0, 0xFF,       0x0,                  []],
    ["Node",    "ILUsed",                "Enumeration", "",  "",         "No",                 ["No",    "Yes"]],
    ["Node",    "NmCAN",                 "Integer",     0,   2,          0,                    []],
    ["Node",    "NmNode",                "Enumeration", "",  "",         "Not",                ["Not",   "Yes"]],
    ["Node",    "NmStationAddress",      "Hex",         0x0, 0xFF,       0x0,                  []],
    ["Node",    "NodeLayerModules",      "String",      "",  "",         "CANoeILNVector.dll", []],
    ["Signal",  "SigType",               "Enumeration", "",  "",         "",                   ["Default", "Range", "RangeSigned", "ASCII", "Discrete", "Control", "ReferencePGN", "DTC", "StringDelimiter", "StringLength", "StringLengthControl"]],
    ["Signal",  "GenSigInactiveValue",   "Integer",     0,   0,          0,                    []],
    ["Signal",  "GenSigSendType",        "Enumeration", "",  "",         "",                   ["cyclic", "OnChange", "OnWrite", "IfActive", "OnChangeWithRepetition", "OnWriteWithRepetition", "IfActiveWithRepetition", "NoSigSendtype"]],
    ["Signal",  "GenSigStartValue",      "Integer",     0,   0,          0,                    []],
    ["Signal",  "GenSigTimeoutValue",    "Integer",     0,   1000000000, 0,                    []],
    ["Signal",  "SPN",                   "Integer",     0,   524287,     0,                    []],
    ["Signal",  "DDI",                   "Integer",     0,   65535,      0,                    []],
    ["Signal",  "SAEDocument",           "String",      "",  "",         "J1939",              []],
    ["Message", "VFrameFormat",          "Enumeration", 0,   8,          "ExtendedCAN",        ["StandardCAN", "ExtendedCAN", "reserved", "J1939PG"]],
    ["Signal",  "SystemSignalLongSymbol","String",      "",  "",         "STRING",             []],
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
            raise ValueError("column number too large: ", str(val))
        return s
    else:
        raise TypeError("column number only support int: ", str(val))


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
                print ("input over range")
        else:
            print ("input invalid")


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
        raise ValueError("Can't find \"Msg Name\" in this sheet")
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
        exit ("ValueError: Duplicate value in Node Colums: {}".format(value))
    # print ("detected nodes: ", template.nodes
    return template


def parse_sig_vals(val_str):
    """
    parse_sig_vals(val_str) -> {valuetable}

    Get signal key:value pairs from string. Returns None if failed.
    """
    sigvals = {}
    if val_str is not None and val_str != '':
        token = re.split('[\;\:\\n]+', val_str.strip())
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
                        print ("waring: ignored signal value definition: " ,token[i], token[i+1])
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


def getint(str, default=None):
    """
    getint(str) -> int
    
    Convert string to int number. If default is given, default value is returned 
    while str is None.
    """
    if str == '':
        if default==None:
            raise ValueError("None type object is unexpected")
        else:
            return default
    else:
        try:
            val = int(str)
            return val
        except (ValueError, TypeError):
            try:
                str = str.replace('0x','')
                val = int(str, 16)
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
        self.nodes = []
        self.messages = []
        self.name = '' ###'CAN'
        self.val_tables = []
        self.version = ''
        self.new_symbols = NEW_SYMBOLS
        self.attr_defs = []
        self.attrs = {}            ###  attrs[namestring] = valuestring
        self._filename = ''
        
        if init:
            self._init_attr_defs()
            for attr in self.attr_defs:
                if attr.name == "DBName":
                    self.attrs["DBName"] = attr.default
                    break

    def _init_attr_defs(self):
        for attr_def in ATTR_DEFS_INIT:
            self.attr_defs.append(CanAttribution(attr_def[1], attr_def[0], attr_def[2], attr_def[3], attr_def[4],
                                                 attr_def[5], attr_def[6]))

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
        line = ["BU_:"]
        if len(self.nodes) > 0:
            for node in self.nodes:
                line.append(node)
            lines.append(' '.join(line) + '\n\n\n')
        else: 
            lines.append(line + '\n\n\n')

        # ! messages
        for msg in self.messages:
            lines.append(str(msg) + '\n\n')
        lines.append('\n\n')

        # ! comments
        lines.append('''CM_ " "''' + ";" + "\n")
        for msg in self.messages:
            comment = msg.comment
            if comment != "":
                line = ["CM_", "BO_", str(msg.msg_id), "\"" + comment + "\"" + ";"]
                lines.append(" ".join(line) + '\n')
            for sig in msg.signals:
                comment = sig.comment
                if comment != "":
                    line = ["CM_", "SG_", str(msg.msg_id), sig.name, "\"" + comment + "\"" + ";"]
                    lines.append(" ".join(line) + '\n')

        # ! attribution defines
        for attr_def in self.attr_defs:
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
            if (val_type == "Enumeration"):
                line.append("ENUM")
                val_range = []
                for val in attr_def.values:
                    val_range.append("\"" + val + "\"")
                line.append(",".join(val_range) + ";")
            elif (val_type == "String"):
                line.append("STRING" + " " + ";")
            elif (val_type == "Hex"):
                line.append("HEX")
                line.append(str(attr_def.min))
                line.append(str(attr_def.max) + ";")
            elif (val_type == "Integer"):
                line.append("INT")
                line.append(str(attr_def.min))
                line.append(str(attr_def.max) + ";")
            lines.append(" ".join(line) + '\n')

        # ! attribution default values
        for attr_def in self.attr_defs:
            line = ["BA_DEF_DEF_"]
            line.append(" \"" + attr_def.name + "\"")
            val_type = attr_def.value_type
            if (val_type == "Enumeration"):
                line.append("\"" + attr_def.default + "\"" + ";")
            elif (val_type == "String"):
                line.append("\"" + attr_def.default + "\"" + ";")
            elif (val_type == "Hex"):
                line.append(str(attr_def.default) + ";")
            elif (val_type == "Integer"):
                line.append(str(attr_def.default) + ";")
            lines.append(" ".join(line) + '\n')

        # ! attribution value of object
        # build-in value "DBName"
        ###line = ["BA_"]
        ###line.append('''"DBName"''')
        ###line.append(self.name + ";")
        ###lines.append(" ".join(line) + '\n')
        for atrdefs in self.attr_defs:
            if atrdefs.object_type == "Network":
                for atr in self.attrs:
                    if atr == atrdefs.name:
                        line = ["BA_"]
                        line.append('\"' + atr + '\"')
                        line.append('\"' + self.attrs[atr] + '\";')
                        lines.append(" ".join(line) + '\n')
            
        
        # ! message attribution values
        for msg in self.messages:
            for attr_def in self.attr_defs:
                #if (msg.attrs.has_key(attr_def.name)):
                if attr_def.name in msg.attrs:
                    if (msg.attrs[attr_def.name] != ''):  ### attr_def.default):
                        line = ["BA_"]
                        line.append("\"" + attr_def.name + "\"")
                        line.append("BO_")
                        line.append(str(msg.msg_id))
                        if attr_def.value_type == "Enumeration":
                            # write enum index instead of enum value
                            line.append(str(attr_def.values.index(str(msg.attrs[attr_def.name]))) + ";")
                        elif attr_def.value_type == "String":
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
                #if sig.init_val is not None: ### and sig.init_val is not 0:
                #    line = ["BA_"]
                #    line.append('''"GenSigStartValue"''')
                #    line.append("SG_")
                #    line.append(str(msg.msg_id))
                #    line.append(sig.name)
                #    line.append(str(sig.init_val))
                #    line.append(';')
                #    lines.append(' '.join(line) + '\n')
                            line = ["BA_"] 
                            line.append('\"' + attr_def.name + '\"')
                            line.append("SG_")
                            line.append(str(msg.msg_id))
                            line.append(sig.name)
                            if (attr_def.value_type == "Enumeration"):
                                line.append(str(attr_def.values.index(str(sig.attrs[attr_def.name]))) + ';')
                            elif (attr_def.value_type == "String"):
                                line.append('\"' + sig.attrs[attr_def.name] + '\";')
                            else:
                                line.append(str(sig.attrs[attr_def.name]) + ';')
                            #line.append(str(sig.attrs[attr]))
                            lines.append(' '.join(line) + '\n')
        # ! Value table define
        for msg in self.messages:
            for sig in msg.signals:
                if sig.values is not None and len(sig.values) >= 1:
                    line = ['VAL_']
                    line.append(str(msg.msg_id))
                    line.append(sig.name)
                    #if debug_enable: print("VALUES: ", end=''); print(sig.values)
                    for key in sig.values:
                        line.append(str(key))
                        line.append('"' + sig.values[key] + '"')
                        #line.append(str(sig.values[key]))
                        #line.append('"'+ key +'"')
                    line.append(';')
                    lines.append(' '.join(line) + '\n')
        return ''.join(lines)

    def add_attr_def(self, name, object_type, value_type, minvalue, maxvalue, default, values=None):
        index = None
        for i in range (0, len(self.attr_defs)):
            if self.attr_defs[i].name == name:
                index = i
                break
        if index is None:
            attr_def = CanAttribution(name, object_type, value_type, minvalue, maxvalue, default, values) 
            self.attr_defs.append(attr_def)
        else:
            print("info: override default attribution definition \'{}\'".format(name))
            self.attr_defs[index].object_type = object_type
            self.attr_defs[index].value_type  = value_type
            self.attr_defs[index].min = minvalue
            self.attr_defs[index].max = maxvalue
            self.attr_defs[index].default = default
            self.attr_defs[index].values = values
            #print(self.attr_defs[index])

    def get_attr_def(self, name):
        ret = None
        for attr_def in self.attr_defs:
            if attr_def.name == name:
                ret = attr_def
                break
        return ret

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

    def set_sig_attr(self, msg_id, sig_name, attr_name, attr_value):
        for msg in self.messages:
            if msg.msg_id == msg_id:
                for sig in msg.signals:
                    if sig.name == sig_name:
                        sig.set_attr(attr_name, attr_value)
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
                attr_def_value_type = attr_def.value_type
                attr_def_values     = [attr for attr in attr_def.values]
                if debug_enable: print(attr_def.name, attr_def_name, attr_def.values, attr_def_values)
                break
        if attr_def_value_type == 'Integer':
            value  = int(value_str)
        elif attr_def_value_type == 'Float':
            value  = float(value_str)
        elif attr_def_value_type == 'String':
            value  = value_str.strip('\"')
        elif attr_def_value_type == 'Enumeration':
            if value_str.isdigit():
                value  = attr_def_values[int(value_str)]
            else:
                value = value_str[1:-1]
        elif attr_def_value_type == 'Hex':
            value  = int(value_str)
        elif attr_def_value_type is None:
            raise ValueError("Undefined attribution definition: {}".format(attr_def_name))
        else:
            raise ValueError("Unkown attribution definition value type: {}".format(attr_def_value_type))
        return value

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
            raise ValueError("Invalid sort option \'{}\'".format(option))

    def load(self, path):
        dbcline = (line.rstrip() for line in open(path, 'r'))  ### generator to read lines from the file helps with multiline comments
        
        for line_rtrimmed in dbcline:
            line_trimmed = line_rtrimmed.lstrip()
            line_split = re.split('[\s\(\)\[\]\|\,\:\;]+', line_trimmed) ### needs better regex to account for optional multiplexor indicator: '', 'M', 'm num', num = 0, 1, 2, ...
            if len(line_split) >= 2:
                # Node
                if line_split[0] == 'BU_':
                    self.nodes.extend(line_split[1:])
                # Message
                elif line_split[0] == 'BO_':
                    msg = CanMessage()
                    msg.msg_id = int(line_split[1])
                    msg.name   = line_split[2]
                    msg.dlc    = int(line_split[3])
                    msg.sender = line_split[4]
                    self.messages.append(msg)
                # Signal
                elif line_split[0] == 'SG_':   ### if SG_, regex "line_trimmed" to break up fields
                    sig = CanSignal()
                    ###                        1=name 2=Multiplexer   3=start_bit 4=len 5=endian 6=sign   7=factor       8=offset             9=min          10=max       11=unit 12=receivers
                    ###                         (1)       (2)             (3)   (4)    (5)   (6)            (7)            (8)                 (9)            (10)          (11)     (12)
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
                        sig.receivers     = re.split('\s+', match.group(12)) # split receivers to list
                    #else:
                    #    sig.name        = line_split[1]
                    #    sig.start_bit   = int(line_split[2])
                    #    sig.sig_len     = int(line_split[3])
                    #    sig.byte_order  = line_split[4][:1]
                    #    sig.value_type  = line_split[4][1:]
                    #    sig.factor      = line_split[5]
                    #    sig.offset      = line_split[6]
                    #    sig.min         = line_split[7]
                    #    sig.max         = line_split[8]
                    #    sig.unit        = line_split[9][1:-1]  # remove quotation markes
                    #    sig.receivers   = line_split[10:]  # receiver is a list
                    msg.signals.append(sig)
                    
                # read in comments from Messages or Signals
                elif line_split[0] == 'CM_':
                    #print(line_trimmed)
                    match0 = re.match(r'CM_\s\"\s*\";', line_trimmed)
                    if match0:
                        #print(line_trimmed)
                        next   # throw away the generic comment
                    match1 = re.match(r'CM_\s+(SG_|BO_)\s+(\d+)\s+([\w\d_]*)\s*\"(.+)\";$', line_trimmed) ### completed comment line ends with ";
                    if match1:
                        end = True
                        comment_type = match1.group(1)
                        message_id = int(match1.group(2))
                        signal_name = match1.group(3)
                        comment_string = match1.group(4).replace("\r", "\n")
                        if debug_enable: print("has end:", str(message_id), signal_name, comment_string)
                        # save comment to message/signal
                        if comment_type == "BO_":
                            self.set_msg_comment(message_id, comment_string)
                        elif comment_type == "SG_":
                            self.set_sig_comment(message_id, signal_name, comment_string)
                    #                            1 = message id, 2 = signal name, 3 = comment content
                    else: 
                        match2 = re.match(r'CM_\s+(SG_|BO_)\s+(\d+)\s+([\w\d_]*)\s*\"(.+)$', line_trimmed) ### start of comment line but no "; to finish it
                        if match2: 
                            end = False
                            comment_type = match2.group(1)
                            message_id = int(match2.group(2))
                            signal_name = match2.group(3)
                            comment_string = match2.group(4)
                            if comment_type == "BO_":
                                self.set_msg_comment(message_id, comment_string.strip(';').strip('\"'))
                            else:
                                if debug_enable: print("no end:", str(message_id), signal_name, comment_string)
                                self.set_sig_comment(message_id, signal_name, comment_string + '\n')
                            while(end==False):
                                line_rtrimmed = next(dbcline)                                                  ### Note - reading next line here
                                if line_rtrimmed == '' or line_rtrimmed == '\n':                               ### blank line in comment
                                    if comment_type == "BO_":
                                        self.append_msg_comment(message_id, '\n')
                                    else: 
                                        if debug_enable: print("cont:...", str(message_id), signal_name, "null line")
                                        self.append_sig_comment(message_id, signal_name, '\n')
                                else: 
                                    match3 = re.match(r'^(?!CM_)(.+)\";$', line_rtrimmed)                      ### continued comment line does not start with CM_ and the "; is detected
                                    if match3:
                                        comment_string = match3.group(1)
                                        end = True
                                        if comment_type == "BO_":
                                            self.append_msg_comment(message_id, comment_string)
                                        else: 
                                            if debug_enable: print("final:  ", str(message_id), signal_name, comment_string)
                                            self.append_sig_comment(message_id, signal_name, comment_string)
                                    else: 
                                        match4 = re.match(r'^(?!CM_)(.+)(?!";)$', line_rtrimmed)               ### continued comment line does not start with CM_ and no "; is detected
                                        if match4:
                                            comment_string = match4.group(1)
                                            if comment_type == "BO_":
                                                self.append_msg_comment(message_id, comment_string)
                                            else: 
                                                if debug_enable: print("cont:  ", str(message_id), signal_name, comment_string)
                                                self.append_sig_comment(message_id, signal_name, comment_string + '\n')
                    
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
                        attr_def_value_type_str = 'Hex'
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
                        raise ValueError("Unkown attribution definition value type: {}".format(attr_def_value_type))

                    self.add_attr_def(attr_def_name, attr_def_object_type, attr_def_value_type_str, \
                                      attr_def_value_min, attr_def_value_max, attr_def_default, attr_def_values)
                # Attribution Definition Default Value
                elif line_split[0] == 'BA_DEF_DEF_':
                    attr_def_name               = line_split[1][1:-1] #remove "
                    attr_def_value_type_str     = None
                    attr_def_default            = self.convert_attr_def_value(attr_def_name, line_split[2])
                    ###print("BA_DEF_DEF: ", end='');print(line_split[2], end='')
                    attr_def                    = self.get_attr_def(attr_def_name)
                    ###print(attr_def)
                    if attr_def is not None:
                        attr_def.default        = attr_def_default
                # Object Attribution definition
                elif line_split[0] == 'BA_':
                    attr_name           = line_split[1][1:-1] # remove "
                    attr_object         = line_split[2]
                    if attr_object == 'BO_':
                        attr_msg_id     = int(line_split[3])
                        ###print("|", line_split[4], "|", sep='', end='\n')
                        attr_value      = self.convert_attr_def_value(attr_name, line_split[4])
                        if attr_value is not None:
                            self.set_msg_attr(attr_msg_id, attr_name, attr_value)
                        else:
                            raise ValueError("Msg \'{}\' attribuition \'{}\' value is {}".format(attr_msg_id, attr_name, line_split[4])) 
                    elif attr_object == 'SG_':
                        attr_msg_id     = int(line_split[3])
                        signal_name     = line_split[4]
                        attr_value      = self.convert_attr_def_value(attr_name, line_split[5])
                        if attr_value is not None: 
                            self.set_sig_attr(attr_msg_id, signal_name, attr_name, attr_value)
                    else: 
                        # network attribute: 
                        attr_name       = line_split[1][1:-1]
                        attr_value      = line_split[2].strip('\"')
                        self.attrs[attr_name] = attr_value
                # Value Tables
                elif line_split[0] == 'VAL_': 
                    quote_split  = re.split('\s*\"\s*', line_trimmed)
                    line_split   = re.split('\s+', quote_split[0])
                    line_split.extend(quote_split[1:-1])
                    val_msg_id   = int(line_split[1])
                    val_sig_name = line_split[2]
                    values = {}
                    for i in range(3, len(line_split)-1, 2):
                        try:
                            values[int(line_split[i])] = line_split[i+1] ### swapped order so that the integer is the key, text is the value
                        except:
                            print(line_split)
                            raise
                    self.set_sig_attr(val_msg_id, val_sig_name, 'values', values)

        dbcline.close()
                        
    def save(self, path=None):
        if (path == None):
            file = open(self._filename + ".dbc", "w")
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
            print ("parse template")
            template = parse_template(sheet)
            #if debug_enable:
                #print (template)
        # ! load network information
        filename = os.path.basename(path).split(".")
        self._filename = ".".join(filename[:-1])
        #self.name = ".".join(filename[:-1]).replace(" ", "_").replace('.', '_').replace('-', '_')  # use filename as default DBName
        self.name = "CAN"
        
        # ! load nodes information
        self.nodes = template.nodes.keys()

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
                        print ("warning: message %s\'s cycle time \"%s\" is invalid, auto set to \'0\'" % (message.name, row_values[template.msg_cycle_col]))
                        msg_cycle = 0
                    message.set_attr("GenMsgCycleTime", msg_cycle)
                    message.set_attr("GenMsgSendType", "cyclic")
                elif (send_type.upper() == "NOMSGSENDTYPE"):
                    message.set_attr("GenMsgSendType", "noMsgSendType")
                else:
                    message.set_attr("GenMsgSendType", "noMsgSendType")
                message.comment = row_values[template.sig_comment_col].replace("\"", "\'").replace(";",",").replace("\r", '\n').replace("\n\n", "\n") # todo
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
                    signal.name = sig_name.replace(' ', '')
                    mux_value = row_values[template.sig_mltplx_col]
                    signal.mux_indicator = mux_value if (mux_value == 'M' or mux_value == '') else 'm' + str(int(mux_value))
                    #print(signal.mux_indicator)
                    signal.start_bit = getint(row_values[template.sig_start_bit_col])
                    signal.sig_len = getint(row_values[template.sig_len_col])
                    
                    byte_order_type = row_values[template.sig_byte_order_col]
                    if (byte_order_type.upper() == "MOTOROLA LSB"):  # todo
                        signal.byte_order = '1'
                    elif (byte_order_type.upper() == "MOTOROLA MSB"):
                        signal.byte_order = '0'
                    else:
                        signal.byte_order = '0'
                        raise ValueError("Unknown signal byte order type: \"%s\""%byte_order_type) # todo: intel
                        
                    if (row_values[template.sig_value_type_col].upper() == "UNSIGNED"):
                        signal.value_type = '+'
                    else:
                        signal.value_type = '-'
                        
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
                            print ("warning: signal %s\'s value table is ignored" % signal.name)
                        else:
                            pass
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
                                print ("warning: message %s\'s sender is set to \"%s\" via signal" %(message.name, nodename))
                            else:
                                print ("warning: message %s\'s sender is conflict to signal \"%s\"" %(message.name, nodename))
                        else:
                            pass
                    if len(signal.receivers) == 0:
                        if len(message.receivers) > 0:
                            signal.receivers = message.receivers
                        else:
                            signal.receivers.append('Vector__XXX')
                    signals.append(signal)
        self.sort();


class CanMessage(object):
    def __init__(self, name='', msg_id=0, dlc=8, sender='Vector__XXX'):
        ''' 
        name: message name
        id: message id (11bit or 29 bit)
        dlc: message data length
        sender: message send node
        '''
        self.name = name
        self.msg_id = msg_id
        self.send_type = ''
        self.dlc = dlc
        self.sender = sender
        self.signals = []
        self.attrs = {}
        self.receivers = []
        self.comment = ''

    def __str__(self):
        para = []
        line = ["BO_", str(self.msg_id), self.name + ":", str(self.dlc), self.sender]
        para.append(" ".join(line))
        for sig in self.signals:
            para.append(str(sig))
        return '\n '.join(para)

    def set_attr(self, name, value):
        self.attrs[name] = value
        
    def set_comment(self, comment):
        self.comment = comment.strip()
        
    def append_comment(self, comment):
        self.comment = self.comment + comment


class CanSignal(object):
    def __init__(self, name='', start_bit=0, sig_len=1, init_val=0):
        self.name = name
        self.mux_indicator = ''
        self.start_bit = start_bit
        self.sig_len = sig_len
        ###self.init_val = init_val
        self.byte_order = '0'
        self.value_type = '+'
        self.factor = 1
        self.offset = 0
        self.min = 0
        self.max = 1
        self.unit = ''
        self.values = {}      ### VAL_
        self.receivers = []
        self.comment = ''     ### CM_ SG_ comment information
        self.attrs = {}      ### BA_ SG_ attribute information
        if init_val != 0:
            self.attrs["GenSigStartValue"] = init_val

    def __str__(self):
        line = ["SG_", self.name + ' '*bool(self.mux_indicator) + self.mux_indicator, ":",
                str(self.start_bit) + "|" + str(self.sig_len) + "@" + self.byte_order + self.value_type,
                "(" + str(self.factor) + "," + str(self.offset) + ")", "[" + str(self.min) + "|" + str(self.max) + "]",
                "\"" + str(self.unit) + "\"", ','.join(self.receivers)]
        return " ".join(line)

    def set_attr(self, name, value):
        if name == 'values':
            self.values = value
        elif name in ["SPN", "DDI", "SigType", "GenSigInactiveValue", "GenSigSendType", "GenSigStartValue", "GenSigTimeoutValue", "SAEDocument", "SystemSignalLongSymbol"]:
            self.attrs[name] = value
        else:
            raise ValueError("Unsupport set attr of \'{}\'".format(name))

    def get_attr(self, name):
        if name == 'values':
            return self.values
        elif name in ["SPN", "DDI", "SigType", "GenSigInactiveValue", "GenSigSendType", "GenSigStartValue", "GenSigTimeoutValue", "SAEDocument", "SystemSignalLongSymbol"]:
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
    def __init__(self, name, object_type, value_type, minvalue, maxvalue, default, values=None):
        self.name = name
        self.object_type = object_type
        self.value_type = value_type
        self.min = minvalue
        self.max = maxvalue
        self.default = default
        self.values = values

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
        return '\n'.join(line)

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

    parse_cmp = subparser.add_parser("cmp", help="Compare difference bettween two dbc files")
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


