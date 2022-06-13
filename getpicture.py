
import docx2txt,os
import cv2

class ExtIMG():
    def __init__(self,path,img_fol):
        self.path = path
        self.img_fol = img_fol
        self.rename = []
        self.clear = []

    def ext_img(self):
        text = docx2txt.process(self.path,self.img_fol)
               
    def check_dup(self):
        count_cor = 0
        proto = cv2.imread('logoYRS.jpg')
        
        f = os.listdir(self.img_fol)
        for img in f:
            comparator = cv2.imread(self.img_fol+img)
            if len(comparator) > 2:
                if sum(comparator[0][count_cor]) == sum(proto[0][count_cor]) :        
                    self.clear.append(img)
                elif sum(comparator[0][count_cor]) != sum(proto[0][count_cor]):
                    self.rename.append(img)
                if count_cor == 3:
                    count_cor = 0
                count_cor += 1
            else :
                self.clear.append(img)

    def rename_remove(self):
        self.name = str(self.path).replace('.docx','').replace('ref/','')
        for index_rename,filenamex in enumerate(self.rename):
            os.rename(self.img_fol+filenamex,'img_fol_ex/'+f'{self.name}_{index_rename}.png')
        for listclear in self.clear:
            os.remove('img_fol/'+listclear)

if __name__ == '__main__':
    M = ExtIMG('')
    M.ext_img()
    M.check_dup()
    M.rename_remove()