from tkinter import *
import random

class MakingAnser:

    def createRanAnswer(self):
        return random.sample(range(1, 10), 3)

class Window:
    components={}
    Number = 0
    def __init__(self):
        self.display = Tk()
        self.display.title('BaseballGame')
        self.display.geometry('400x300')
        Operation.CreateAnswer()
        Window.components= self.contentsplace()

    def contentsplace(self):
        aLabel = Label(self.display,text ="Please enter 3-digits")
        aLabel.pack()

        Answer = Entry(self.display)
        Answer.bind('<Return>',Operation.compareAnswers)
        Answer.pack()
        Answer.focus_set()

        bLabel = Label(self.display)
        bLabel.pack()

        cLabel = Label(self.display)
        cLabel.pack()

        return {'DiLabel':aLabel,'field':Answer,'AnLabel': bLabel,'reLabel':cLabel}

class Operation:
    number =0
    checkResult = {'strike': 0, 'ball': 0}
    flag = 0
    @classmethod
    def CreateAnswer(cls):
        Operation.number = random.sample(range(1,10),3)

    @classmethod
    def compareAnswers(cls,event):
        Operation.checkResult = {'strike': 0, 'ball': 0}

        field:Entry= Window.components['field']
        field.focus_set()
        Value = field.get()
        if len(Value) != 3:
            Operation.checkResult = "Please enter 3-digits"
            Operation.flag = 1
        else:
            uValue = [int(int(Value)/100),int((int(Value)/10)%10) ,int(Value)%10]

            print(uValue,Operation.number)

            for position in range(len(Operation.number)):
                if uValue[position] == Operation.number[position]:
                    Operation.checkResult['strike'] += 1
                elif uValue[position] in Operation.number:
                    Operation.checkResult['ball'] += 1
        if Operation.flag == 0:
            Operation.Update()
    @classmethod
    def Update(cls):
        AnLabel:Label = Window.components['AnLabel']
        AnLabel.configure(text = Operation.checkResult)

        if Operation.checkResult['strike'] == 3:
            ReLabel : Label =Window.components['reLabel']
            ReLabel.configure(text = 'Finish and Enter New number')
            Operation.CreateAnswer()
        else :
            ReLabel: Label = Window.components['reLabel']
            ReLabel.configure(text='')


app = Window()
app.display.mainloop()