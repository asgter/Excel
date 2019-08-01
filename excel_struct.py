import xlrd

class ExcelModel():
    def __init__(self,filepath):
        self.x1 = xlrd.open_workbook(filepath)
        self.excel_dict = {}

    def construct(self):
        pass

    def get_sheet(self):
        for i in self.x1.sheet_names():
            self.excel_dict[i] = self.x1.sheet_by_name(i)

    def walk_cells(self,sheet):
        self.node_warehouse = []
        node_line = {}
        node_dict = {}
        for i in range(1,sheet.nrows+1):
            node_line = {i:None}
            for j in range(1,sheet.ncols+1):
                node_dict = {j:None}
                val = sheet.cell(i-1,j-1).value
                # print(i,j,val)
                # print(i, j,bool(len(val) if type(val) == str else val > 0))
                if j == 1 and (len(val) if type(val) == str else int(val)+1):
                    cell_node = Node(val,i,j,3)
                    node_dict[j] = cell_node
                    node_line[i] = node_dict
                elif j != 1 and (len(val) if type(val) == str else int(val)+1):
                    if node_line.get(i):
                        cell_node = Node(val,i,j,4)
                        node_dict[j] = cell_node
                        node_line[i].update(node_dict)
                    else:
                        cell_node = Node(val,i,j,3)
                        node_dict[j] = cell_node
                        node_line[i] = node_dict
                else:
                    if not (len(val) if type(val) == str else int(val)+1):
                        continue
                    else:
                        cell_node = Node(val, i, j, 4)
                        node_dict[j] = cell_node
                        node_line[i].update(node_dict)
            if node_line[i] is None:
                continue
            else:
                self.node_warehouse.append(node_line)
        # print(self.node_warehouse)
        #
        # print(node_line.get(9).get(1).cell,node_line.get(9).get(1).row,node_line.get(9).get(1).col)

        self.distribute()
        # for n in self.node_warehouse:
        #     for i in n.values():
        #         print('\n')
        #         for j in i.values():
        #             try:
        #                 print(j.right_child.cell,end=" ")
        #             except AttributeError as e:
        #                 pass

        # print(self.node_warehouse[0])



    def distribute(self):
        for n in range(1,len(self.node_warehouse)+1):
            try:
                for k,v in self.node_warehouse[n].items():
                    if len(v)> 1:
                        for i in range(1,len(v)):
                            v[i].right_child = v[i+1]
                    else:
                        v.right_child = None
            except :
                pass


    def find_node(self,row,col):
        for i in range(1,len(self.node_warehouse)+1):
            try:
                if self.node_warehouse[i][row][col]:
                    return self.node_warehouse[i][row][col]
                else:
                    continue
            except :
                continue
        return Node("此单元格不存在",1,1,0)

    def traver_line(self,node,list):
        list = list
        if  (node.getattr("right_child") if type(node) == "Node" else bool(node)):
            self.traver_line(node.right_child,list)
            list.append(node.cell)
        else:
            try:
                list.append(node.cell)
            except AttributeError:
                pass
        return list[::-1]

    def find_row(self,row):
        i = 1
        list = []
        for i in range(1,10):
            try:
                node = self.find_node(row,i)
                break
            except:
                continue
        # print(node)
        list1 = self.traver_line(node, list)
        print(list1)
        return list1









class Sheet():
    def __init__(self):
        pass


class Node():
    def __init__(self,cell,row,col,label):
        if type(cell) == str:
            self.cell = cell.strip(" ")
        else:
            self.cell = cell
        self.row = row
        self.col = col
        # self.left_child = None
        self.right_child = None
        self.label = label

    def __getattr__(self, item):
        # print('----> from getattr:你找的属性不存在')
        return False



if __name__ == "__main__":
    em = ExcelModel(r"C:\Users\Draven\Desktop\123.xlsx")
    em.get_sheet()
    # print(em.excel_dict)
    em.walk_cells(em.excel_dict["Sheet1"])
    # print(em.node_warehouse)
    content = em.find_node(13,3)
    # print(content)
    em.find_row(18)