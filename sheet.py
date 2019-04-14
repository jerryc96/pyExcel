from openpyxl import load_workbook
class Sheet:
    def __init__(self, sheet):
        # matrix for holding excel data
        self.data = []
        self.keys = []
        ind = 0
        key_row = True
        for row in sheet.rows:
            if key_row:
                key_row = False
                for cell in row:
                    self.keys.append(str(cell.value))                
            else:
                self.data.append([])
                for cell in row:
                    self.data[ind].append(str(cell.value))
                ind += 1
    '''
    lookup(row, column)
    given an excel key field, lookup and return the value held at that field in a list.
    
    options:
    [str, str] (key, value): input a key, value pair, returns the record(s) that matches
    the value for that key field.
    '''
    def lookup(self, key, value):
        result = []
        if key not in self.keys:
            raise KeyError('No' + ' field named: ' + key)
        else:
            keyInd = self.keys.index(key)
            for row in self.data:
                if value == row[keyInd]:
                    result.append(row)
            if len(result) == 0:
                raise KeyError(value + ' does not exist')
            return result;
    
    # id is unique, so there's always 1 per query
    def lookup_ID(self, netID):
        return self.lookup(self.keys[0], netID)[0]