from xlrd import open_workbook
import matplotlib.pyplot as plt
import networkx as nx
import random


data = dict()
docs = {}
exam_teacher = {}
color_list = []
timeAdd_exam = []

def read_xls(wb):  
    for s in wb.sheets():
        for row in range(2,s.nrows):
            values = []
            for col in range(s.ncols):                            
                values.append(s.cell(row, col).value)
            exam = values[0]
            if exam not in docs:
                docs[exam] = []
            docs[exam].append(' '.join([values[1], values[2]]))
            exam_teacher[exam]= values[4]
            if exam not in data:
                data[exam] = []
            data[exam].append(('_'.join([values[1], values[2]]), values[3], values[4])) 
            global timeAdd_exam
            if values[5] != '':
                timeAdd_exam.append(exam)
        break
    timeAdd_exam = list(set(timeAdd_exam))
    return docs.keys()

def build_graph(nodes):
    G = nx.Graph()
    G.add_nodes_from(nodes)
    total_edges = [(x, y) for i, x in enumerate(nodes) for y in nodes[i+1:]]
    for node1, node2  in total_edges:
        if list(set(docs[node1]) & set(docs[node2])):
            G.add_edge(node1, node2)        
    return G
              
def coloring(G,nodes): 
    global timeAdd_exam
    nodes_color = {}
    for node in nodes:
        nodes_color[node]=0        
    for node in nodes:
        neighbor_node=nx.all_neighbors(G,node)
        neighbor_node_colors= [nodes_color[node1] for node1 in neighbor_node]
        for k in range(1,10000):
            if k not in neighbor_node_colors:
                nodes_color[node]=k 
                break              
    color_list = [nodes_color[node] for node in nodes]
    return nodes_color
    
def drawing(G,nodes,color_list):      
    # draw graph
    pos = nx.shell_layout(G)
    nx.draw(G,pos,nodelist=nodes,node_size=500,node_color=color_list,width=0.2,lable='2.xcl')
    #show graph
    plt.show()   
        
def Group(nodes_color):
    color_group = {}  
    for k, v in nodes_color.items():
        apart = []
        apart_num = []      
        for s in data[k]:
            apart.append(s[1])
            teacher = s[2] 
        for s in apart:
            apart_num.append((apart.count(s), s))
        apart_num = list(set(apart_num))
        apart_num = '('+','.join([''.join([str(x), y]) for x, y in apart_num])+')'
        stu_num = len(data[k])
     
        if v not in color_group:
            color_group[v] = []
        if [k, apart_num,  str(stu_num), teacher] not in color_group[v]:
                color_group[v].append([k, apart_num,  str(stu_num), teacher])
    return color_group

def Ordering(color_group):
    order_group = {}
    for j in color_group.keys():
        order_group[j] = []   
          
    Part_max = {}
    for k, v in color_group.items():
        Part_max[k] = max([int(s[2]) for s in v])
    for i, color in enumerate(sorted(color_group.keys(), key=lambda d: Part_max[d], reverse=True)):
        order_group[i+1] = color_group[color]                  
    return order_group   

def find(order_group,timeAdd_list):
    timeAddExam_color = {}
    for exam in timeAdd_list:
        for k,v in order_group.items():
            for s in v:
                if s[0] == exam:
                    timeAddExam_color[exam] = k
    return timeAddExam_color
                    
def teacher_group(order_group):
    teacher_group  = {}
    for k, v in order_group.items() :
        teacher_group[k] = [s[3] for s in v]
    return teacher_group
        
def match(timeAddExam_color1,timeAddExam_color2,teacher_group1, teacher_group2):
    new_group = {}
    #print timeAddExam_color1,timeAddExam_color2
    for s,t in timeAddExam_color1.items():
        for m,n in timeAddExam_color2.items():
            new_group[t] = n
            del teacher_group1[t]
            del teacher_group2[n]
            
    for k, v in teacher_group1.items():
        for m, n in teacher_group2.items():
            if list(set(v) & set(n)):
                new_group[k] = m
                del teacher_group2[m]
                break
    return new_group             
            
               
def main(fname1,fname2):  
    global timeAdd_exam
    wb1 = open_workbook(fname1)
    nodes1 = read_xls(wb1)
    G1 = build_graph(nodes1)
    min_color_num1 = 10000
    min_nodes1 = nodes1[:]
    for _ in range(100):
        random.shuffle(nodes1)
        color_num1 = len(set(coloring(G1,nodes1).values()))
        if color_num1 < min_color_num1:
            min_color_num1 = color_num1
            min_nodes1 = nodes1[:]

    nodes1 = min_nodes1[:]
    nodes_color1 = coloring(G1,nodes1)
    
    #nodes_color1 = nodes_color[:]
    color_group1 = Group(nodes_color1)
    order_group1 = Ordering(color_group1)
    exam_stu1 = docs.copy()
    timeAdd_exam1 = timeAdd_exam[:]
    timeAddExam_color1 = find(order_group1,timeAdd_exam1)
    
    timeAdd_exam = []
    data.clear()
    docs.clear()
    
    wb2 = open_workbook(fname2)
    nodes2 = read_xls(wb2)
    G2 = build_graph(nodes2)
    min_color_num2 = 10000
    min_nodes2 = nodes2[:]
    for _ in range(100):
        random.shuffle(nodes2)
        color_num2 = len(set(coloring(G2,nodes2).values()))
        if color_num2 < min_color_num2:
            min_color_num2 = color_num2
            min_nodes2 = nodes2[:]
   
    nodes2 = min_nodes2[:]
    nodes_color2 = coloring(G2,nodes2)
    color_group2 = Group(nodes_color2)
    order_group2 = Ordering(color_group2)
    
    timeAdd_exam2 = timeAdd_exam[:]
    timeAddExam_color2 = find(order_group2,timeAdd_exam2)
    exam_stu2 = docs.copy()
    
    teacher_group1 = teacher_group(order_group1)
    teacher_group2 = teacher_group(order_group2)
    
    Mapping = match(timeAddExam_color1,timeAddExam_color2,teacher_group1, teacher_group2)
    

    remain_order_group2 = range(1, len(order_group2.keys()) + 1)
    remain_order_group2 = [color for color in remain_order_group2 if color not in Mapping.values()]
    
    for color in range(1, len(order_group1.keys()) +1):
        if color not in Mapping.keys():
            try:
                Mapping[color] = remain_order_group2[0]
                remain_order_group2.remove(remain_order_group2[0])
            except IndexError, e:
                break
          
    return order_group1,order_group2,Mapping,exam_stu1,exam_stu2,timeAdd_exam1,timeAdd_exam2
