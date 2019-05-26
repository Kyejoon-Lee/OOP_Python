from tkinter import *
from pygame import *
import pandas as pd
import random
import xlsxwriter
from datetime import datetime

class WindowContents:
    components = {}

    def __init__(self, display, controller):
        self.controller = controller
        self.questgenerater = Question()
        self.quest, self.answer, self.ans_index = self.questgenerater.question_generater()
        self.usersheet = [] # Answersheet from player
        self.score = 0
        self.display = display

        WindowContents.components = self.placingComponents()

        Label(display, name="qnum", text="#%s"%(len(self.usersheet)+1)).place(x=0,y=0)

        #relief options : flat, groove, raised, ridge, solid, or sunken
        Label(display, name='aLabel', text="%s" % (self.quest[len(self.usersheet)]),
              font=("Helvetica Bold", 36), relief='solid', height = 2, width = 25).pack()

        # warning message
        Label(display, name='aLabel_1', text="", fg="Red", font=("Helvetica Bold", 15)).place(x=220, y=420)

        # setting radio button
        self.selectedRadio = IntVar()
        List = self.answer[len(self.usersheet)]
        Radiobutton(display, name='aRadioLabel_1', text=List[0], font=("Helvetica Bold", 24), value=0,
                    variable=self.selectedRadio).place(x=20, y=140)
        Radiobutton(display, name='aRadioLabel_2', text=List[1], font=("Helvetica Bold", 24), value=1,
                    variable=self.selectedRadio).place(x=20, y=195)
        Radiobutton(display, name='aRadioLabel_3', text=List[2], font=("Helvetica Bold", 24), value=2,
                    variable=self.selectedRadio).place(x=20, y=250)
        Radiobutton(display, name='aRadioLabel_4', text=List[3], font=("Helvetica Bold", 24), value=3,
                    variable=self.selectedRadio).place(x=20, y=305)
        Radiobutton(display, name='aRadioLabel_5', text=List[4], font=("Helvetica Bold", 24), value=4,
        variable = self.selectedRadio).place(x=20, y=360)

        # setting confirm button
        Button(display, name='confirmButton', text="Confirm",
               command=self.confirmbuttonPressed).place(x=400, y=420, width=170, height=50)

        # setting back button
        Button(display, name='backButton', text="Back", command=self.backbuttonPressed).place(
            x=30, y=420, width=170, height=50)

        # to make Radiobutton none
        self.selectedRadio.set(None)

    def placingComponents(self):
        userinfo = Label(self.display,text="user name")
        userinfo.pack()

        return {'info_label': userinfo}

    # if you press confirm button
    def confirmbuttonPressed(self):

        # Play button sound
        mixer.init()
        mixer.music.load('shortbutton.wav')
        mixer.music.play()

        while mixer.music.get_busy():
            time.Clock().tick(3)

        app: MyWindow = MyWindow.components['app']
        nLabel = app.display.nametowidget('aLabel')
        RadioLabel_1 = app.display.nametowidget('aRadioLabel_1')
        RadioLabel_2 = app.display.nametowidget('aRadioLabel_2')
        RadioLabel_3 = app.display.nametowidget('aRadioLabel_3')
        RadioLabel_4 = app.display.nametowidget('aRadioLabel_4')
        RadioLabel_5 = app.display.nametowidget('aRadioLabel_5')
        nLabel_1 = app.display.nametowidget('aLabel_1')
        nLabel_1.configure(text='')

        qnumLabel = app.display.nametowidget('qnum')

        try:
            # Get player's answer
            UserAnswer = int(self.selectedRadio.get())
        except:
            # If player didn't check any radiobutton
            nLabel_1.configure(text='check the answer')

        else:
            self.usersheet.append(UserAnswer)
            # if player solve 10 probelms
            if len(self.usersheet) >= 10:
                ScorePage.checklist= []
                self.score = self.questgenerater.Score(self.usersheet)
                #self.score = (점수, 오답노트)
                # print(self.score[0])
                ScorePage.score = self.score[0]
                # print(ScorePage.score)
                ScorePage.checklist.append(self.score)
                #print('!!!',ScorePage.checklist)

                #For update we initialize the all Question and usersheet
                self.questgenerater = Question()
                self.quest, self.answer, self.ans_index = self.questgenerater.question_generater()
                self.usersheet = []
                NumofAnswer = len(self.usersheet)
                qnumLabel.configure(text="#%s"%str(NumofAnswer+1))
                nLabel.configure(text='%s'%self.quest[0])
                RadioLabel_1.configure(text=self.answer[0][0])
                RadioLabel_2.configure(text=self.answer[0][1])
                RadioLabel_3.configure(text=self.answer[0][2])
                RadioLabel_4.configure(text=self.answer[0][3])
                RadioLabel_5.configure(text=self.answer[0][4])
                scoreLabel: Label = ScorePage.components['label_L']
                scoreLabel.configure(text="Your score is %d!"%ScorePage.score, fg="Red")

                scorebLabel: Label = ScorePage.components['qlabel']

                # Set Text Color Red if the Answer is Wrong
                if (ScorePage.checklist[0][1][9][1] == ScorePage.checklist[0][1][9][2]):
                    scorebLabel.configure(text="10. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][9][0], ScorePage.checklist[0][1][9][1], ScorePage.checklist[0][1][9][2]), fg="Black")

                elif (ScorePage.checklist[0][1][9][1] != ScorePage.checklist[0][1][9][2]):
                    scorebLabel.configure( text="10. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][9][0], ScorePage.checklist[0][1][9][1], ScorePage.checklist[0][1][9][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['wlabel']

                if (ScorePage.checklist[0][1][8][1] == ScorePage.checklist[0][1][8][2]):
                    scorebLabel.configure(text="9. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][8][0], ScorePage.checklist[0][1][8][1], ScorePage.checklist[0][1][8][2]), fg="Black")

                elif (ScorePage.checklist[0][1][8][1] != ScorePage.checklist[0][1][8][2]):
                    scorebLabel.configure(text="9. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][8][0], ScorePage.checklist[0][1][8][1], ScorePage.checklist[0][1][8][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['elabel']

                if (ScorePage.checklist[0][1][7][1] == ScorePage.checklist[0][1][7][2]):
                    scorebLabel.configure(text="8. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][7][0], ScorePage.checklist[0][1][7][1], ScorePage.checklist[0][1][7][2]), fg="Black")

                elif (ScorePage.checklist[0][1][7][1] != ScorePage.checklist[0][1][7][2]):
                    scorebLabel.configure(text="8. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][7][0], ScorePage.checklist[0][1][7][1], ScorePage.checklist[0][1][7][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['rlabel']

                if (ScorePage.checklist[0][1][6][1] == ScorePage.checklist[0][1][6][2]):
                    scorebLabel.configure(text="7. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][6][0], ScorePage.checklist[0][1][6][1], ScorePage.checklist[0][1][6][2]), fg="Black")

                elif (ScorePage.checklist[0][1][6][1] != ScorePage.checklist[0][1][6][2]):
                    scorebLabel.configure(text="7. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][6][0], ScorePage.checklist[0][1][6][1], ScorePage.checklist[0][1][6][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['tlabel']

                if (ScorePage.checklist[0][1][5][1] == ScorePage.checklist[0][1][5][2]):
                    scorebLabel.configure(text="6. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][5][0], ScorePage.checklist[0][1][5][1], ScorePage.checklist[0][1][5][2]), fg="Black")

                elif (ScorePage.checklist[0][1][5][1] != ScorePage.checklist[0][1][5][2]):
                    scorebLabel.configure(text="6. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][5][0], ScorePage.checklist[0][1][5][1], ScorePage.checklist[0][1][5][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['ylabel']

                if (ScorePage.checklist[0][1][4][1] == ScorePage.checklist[0][1][4][2]):
                    scorebLabel.configure(text="5. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][4][0], ScorePage.checklist[0][1][4][1], ScorePage.checklist[0][1][4][2]), fg="Black")

                elif (ScorePage.checklist[0][1][4][1] != ScorePage.checklist[0][1][4][2]):
                    scorebLabel.configure(
                        text="5. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][4][0], ScorePage.checklist[0][1][4][1], ScorePage.checklist[0][1][4][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['ulabel']

                if (ScorePage.checklist[0][1][3][1] == ScorePage.checklist[0][1][3][2]):
                    scorebLabel.configure(
                        text="4. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][3][0], ScorePage.checklist[0][1][3][1], ScorePage.checklist[0][1][3][2]), fg="Black")

                elif (ScorePage.checklist[0][1][3][1] != ScorePage.checklist[0][1][6][2]):
                    scorebLabel.configure(
                        text="4. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][3][0], ScorePage.checklist[0][1][3][1], ScorePage.checklist[0][1][3][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['ilabel']

                if (ScorePage.checklist[0][1][2][1] == ScorePage.checklist[0][1][2][2]):
                    scorebLabel.configure(text="3. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][2][0], ScorePage.checklist[0][1][2][1], ScorePage.checklist[0][1][2][2]), fg="Black")

                elif (ScorePage.checklist[0][1][2][1] != ScorePage.checklist[0][1][2][2]):
                    scorebLabel.configure( text="3. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][2][0], ScorePage.checklist[0][1][2][1], ScorePage.checklist[0][1][2][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['olabel']

                if (ScorePage.checklist[0][1][1][1] == ScorePage.checklist[0][1][1][2]):
                    scorebLabel.configure( text="2. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][1][0], ScorePage.checklist[0][1][1][1], ScorePage.checklist[0][1][1][2]), fg="Black")

                elif (ScorePage.checklist[0][1][1][1] != ScorePage.checklist[0][1][1][2]):
                    scorebLabel.configure(text="2. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][1][0], ScorePage.checklist[0][1][1][1], ScorePage.checklist[0][1][1][2]), fg="Red")

                scorebLabel: Label = ScorePage.components['plabel']

                if (ScorePage.checklist[0][1][0][1] == ScorePage.checklist[0][1][0][2]):
                    scorebLabel.configure(text="1. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][0][0], ScorePage.checklist[0][1][0][1], ScorePage.checklist[0][1][0][2]), fg="Black")

                elif (ScorePage.checklist[0][1][0][1] != ScorePage.checklist[0][1][0][2]):
                    scorebLabel.configure(text="1. The answer of {} is {} and you choosed {}".format(ScorePage.checklist[0][1][0][0], ScorePage.checklist[0][1][0][1], ScorePage.checklist[0][1][0][2]), fg="Red")

                self.selectedRadio.set(None)

                self.controller.show_frame("ScorePage")

            # Make next question
            else:
                NumofAnswer = len(self.usersheet)
                nextquestion = self.quest[NumofAnswer]
                NewAnswer = self.answer[NumofAnswer]
                self.selectedRadio.set(None)
                qnumLabel.configure(text="#%s"%str(NumofAnswer+1))
                nLabel.configure(text=nextquestion)

                RadioLabel_1.configure(text=NewAnswer[0])
                RadioLabel_2.configure(text=NewAnswer[1])
                RadioLabel_3.configure(text=NewAnswer[2])
                RadioLabel_4.configure(text=NewAnswer[3])
                RadioLabel_5.configure(text=NewAnswer[4])

    # if player press back button
    def backbuttonPressed(self):

        app: MyWindow = MyWindow.components['app']
        nLabel = app.display.nametowidget('aLabel')
        nLabel_1 = app.display.nametowidget('aLabel_1')
        qnumLabel = app.display.nametowidget('qnum')

        nLabel_1.configure(text='')

        # if player press back button at first problem
        if len(self.usersheet) == 0:
           self.controller.show_frame("StartPage")

        else :
            # delete answer of previous question and display problem again
            RadioLabel_1 = app.display.nametowidget('aRadioLabel_1')
            RadioLabel_2 = app.display.nametowidget('aRadioLabel_2')
            RadioLabel_3 = app.display.nametowidget('aRadioLabel_3')
            RadioLabel_4 = app.display.nametowidget('aRadioLabel_4')
            RadioLabel_5 = app.display.nametowidget('aRadioLabel_5')

            self.usersheet.pop(-1)
            NumofAnswer = len(self.usersheet)
            Backquestion = self.quest[NumofAnswer]
            self.selectedRadio.set(None)
            BackAnswer = self.answer[NumofAnswer]
            qnumLabel.configure(text="#%s" % str(NumofAnswer + 1))

            nLabel.configure(text=Backquestion)
            RadioLabel_1.configure(text=BackAnswer[0])
            RadioLabel_2.configure(text=BackAnswer[1])
            RadioLabel_3.configure(text=BackAnswer[2])
            RadioLabel_4.configure(text=BackAnswer[3])
            RadioLabel_5.configure(text=BackAnswer[4])

# 문제 출제 코드
class Question:
    def __init__(self):
        # 엑셀 파일 불러오기
        fileName = 'VocaTest.xlsx'
        dataFile = pd.ExcelFile(fileName)
        # words = DataFrame
        words = dataFile.parse()

        # DataFrame의 NaN값을 '-'으로 변경
        for i in range(2,5):
            del_col = "Synonym"+str(i)
            words[del_col] = words[del_col].fillna('-')

        # DataFrame 단어리스트
        # print(words)

        # DataFrame을 List로 변경
        # [TargetWord, Syn1, Syn2, Syn3, Syn4]
        # aDict = {}
        self.aList = []
        for index, row in words.iterrows():
                self.aList.append([row['TargetWord'],row['Synonym1'],str(row['Synonym2']),str(row['Synonym3']),str(row['Synonym4'])])
                #aDict[row['TargetWord']] = {'Synonym1':row['Synonym1'], 'Synonym2':row['Synonym2'], 'Synonym3':row['Synonym3'], 'Synonym4':row['Synonym4']}

        # print(aDict)
        # print(self.aList)

        # List 길이 : 46
        # print(len(aList))

        self.questionlist = [] # 문제 list
        self.selectlist = [] # 선택지 list
        self.ans_index = []

    # 문제 출제 함수
    def question_generater(self):

        # 초기화
        questionindex = []
        self.questionlist = []
        self.selectlist = []

        # 랜덤으로 10개 문제 출제
        while len(questionindex)<10:
            # 0~44 범위의 수 랜덤 생성
            i = random.randint(0, len(self.aList)-1)

            # 중복 검사
            if i not in questionindex:
                questionindex.append(i)

        # 문제로 출제할 list by index
        # print("questionindex",questionindex)

        # 문제 출제
        for i in questionindex:
            #문제 하나가 들어갈 임시 list
            question_temp = []

            # aList에서 i번째 단어를 문제로 출제
            self.questionlist.append(self.aList[i][0])

            # 정답이 될 수 있는 Synonum list
            right_syn = self.aList[i][1:]
            right_ans = '-'

            # '-'이 아닌 값이 선택될 까지 Synonum list에서 랜덤으로 선택
            while right_ans == '-':
                right_ans = random.choice(right_syn)

            question_temp.append(right_ans)

            # 오답지 저장 list
            error_list = []

            # 오답 Synonum 생성을 위한 error 단어 list by index
            while len(error_list)<4:
                # err 범위 : 0 ~ 44 (len이 46)
                err = random.randint(0, len(self.aList) - 1)

                # question으로 출제되는 단어의 index와 겹치지 않도록
                if err != i and err not in error_list:
                    error_list.append(err)

            # 오답 Synonum 생성
            for index in error_list:
                # 오답지용으로 선택한 단어의 synonym list
                select_syn = self.aList[index][1:]

                err_word='-'

                # 오답지용으로 선택한 단어의 synonym 중 하나 선택
                while err_word == '-':
                    err_word = random.choice(select_syn)

                    # 동일 답지 방지
                    if err_word in question_temp:
                        err_word = '-'

                question_temp.append(err_word)

            # print("question_temp", question_temp)

            # 최종 문제지 list에 n번째 문제 추가
            self.selectlist.append(question_temp)

            # print("selectlist",self.selectlist)

        self.ans_index = []

        for wordlist in self.selectlist:
            # 실제 답은 list의 0번째 요소
            basic_ans = wordlist[0]

            # 답지 Shuffle
            random.shuffle(wordlist)

            # self.ans_index : 정답의 변경된 index
            self.ans_index.append(wordlist.index(basic_ans))

        # 문제 출제 완료 확인
        # for i in range(0,10):
        #     print(self.questionlist[i])
        #     print(self.selectlist[i])
        #     print(self.ans_index[i])

        return self.questionlist, self.selectlist, self.ans_index

    # 점수 및 오답노트를 위한 Data 생성
    def Score(self, user_ans):
        score = 0
        checklist_notebook = []

        # 0~9
        for i in range(0,len(user_ans)):
            # 디버그용 확인
            # print(self.selectlist[i][self.ans_index[i]], self.selectlist[i][user_ans[i]])
            # 원래 답지
            # print(self.ans_index)

            # 한 문제 당 10점
            if user_ans[i] == self.ans_index[i]:
                score += 10
                checklist_notebook.append([self.questionlist[i], self.selectlist[i][self.ans_index[i]], self.selectlist[i][user_ans[i]]])

            else:
                # [문제, 답]
                checklist_notebook.append([self.questionlist[i], self.selectlist[i][self.ans_index[i]], self.selectlist[i][user_ans[i]]])
                # print("The answer of",self.questionlist[i],"is",self.selectlist[i][self.ans_index[i]], "but you choosed", self.selectlist[i][user_ans[i]])
                #checklist_notebook

        return score, checklist_notebook


class MyWindow:
    components = {}

    def __init__(self, size='600x500'):
        self.display= Tk()
        self.display.title("VOCA QUIZ")
        self.display.geometry(size)

        #We make a Frame for window change
        container = Frame(self.display)
        container.pack(side="bottom", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (StartPage, PageOne,ScorePage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # 모든 페이지를 같은 위치로 설정
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")
        self.display.mainloop()

    # 페이지 이동
    def show_frame(self, page_name):
        # 첫 번째 페이지
        if page_name=="PageOne":
            name_entry: Entry = StartPage.components['name_entry']
            user_name = name_entry.get()

            stid_entry: Entry = StartPage.components['stid_entry']
            student_id = stid_entry.get()

            user_info : Label = WindowContents.components['info_label']
            user_info.configure(text="%s %s"%(student_id, user_name))

        # 마지막 페이지
        elif page_name == "ScorePage" :
            name_entry: Entry = StartPage.components['name_entry']
            user_name = name_entry.get()

            stid_entry: Entry = StartPage.components['stid_entry']
            student_id = stid_entry.get()

            user_info : Label = ScorePage.components['info_label']
            user_info.configure(text="%s %s"%(student_id, user_name))

        frame = self.frames[page_name]
        # 창을 앞으로 가져옴
        frame.tkraise()

class StartPage(Frame):
    components = {}

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.display = parent
        self.controller = controller

        StartPage.components = self.placingComponents()

    def placingComponents(self):
        controller = self.controller

        alabel = Label(self, text="VOCA QUIZ : SYNONYM", font=("Helvetica Bold", 36), height=4, anchor=CENTER)
        alabel.pack()

        intro = Label(self, text="Please Enter Your Personal Information!", font=("Helvetica", 15), anchor=CENTER)
        intro.place(x=130,y=150)

        namelabel = Label(self, text="이름")
        namelabel.place(x=200, y=202)

        nameentry = Entry(self, justify=CENTER)
        nameentry.place(x=240, y=200)

        numlabel = Label(self, text='학번')
        numlabel.place(x=200, y=242)

        StNum = Entry(self, justify=CENTER)
        StNum.place(x=240, y=240)

        Zbutton = Button(self, text='Show Production Crew', command=self.productioncrew)
        Zbutton.place(x=400, y=420, width=170, height=50)

        button = Button(self, text='start', command=lambda: controller.show_frame("PageOne"))
        button.place(x=400, y=360, width=170, height=50)

        crewlabel = Label(self, text="")
        crewlabel.place(x=0,y=0)

        return {'title_label': alabel,'intro_label': intro, 'name_label': namelabel,'name_entry':nameentry,'num_label':numlabel,'stid_entry':StNum,'crew_button': Zbutton,'start_button':button,'crew_label': crewlabel}

    # Crew Introduction
    def productioncrew(self):

        crew_label : Label = StartPage.components['crew_label']
        crew_button : Button = StartPage.components['crew_button']

        # 버튼을 누르면 글씨가 생성되었다 사라짐
        if len(crew_label['text'])==0:
            crew_label.configure(text = "MADE BY\n2013314264 이계준\n2013311822 김경회\n2014312937 양윤석\n2015318802 최주아\n2018311658 차민지")
            crew_label.place(x=250, y=360)
            crew_button.configure(text="Hide Production Crew")

        else:
            crew_button['text']="Show Production Crew"
            crew_label['text'] = ""

# 문제 풀이 페이지
class PageOne(Frame):
    def __init__(self, parent, controller):
        self.display = self
        Frame.__init__(self, parent)
        self.controller = controller
        MyWindow.components = self.New()
    #This call the WindowContents class
    def New(self):
        contents = WindowContents(self.display, self.controller)
        return {'app': self, 'contents': contents}

# 마지막 페이지
class ScorePage(Frame):
    components = {}
    score = 0
    checklist = []

    def __init__(self,parent,controller):
        Frame.__init__(self, parent)
        self.controller = controller

        ScorePage.components = self.placingComponents()

    #placing the components
    def placingComponents(self):
        controller = self.controller

        userinfo = Label(self,text="")
        userinfo.pack()

        label = Label(self,text ="Your Score is %d!"%ScorePage.score, relief = 'solid', font=("Helvetica Bold", 36), width=25, height = 2)
        label.pack()

        # 오답노트용 Label
        qlabel = Label(self,text="%s" %ScorePage.checklist)
        qlabel.place(x=30,y=375)
        wlabel = Label(self, text="%s" % ScorePage.checklist)
        wlabel.place(x=30, y=350)
        elabel = Label(self, text="%s" % ScorePage.checklist)
        elabel.place(x=30, y=325)
        rlabel = Label(self, text="%s" % ScorePage.checklist)
        rlabel.place(x=30, y=300)
        tlabel = Label(self, text="%s" % ScorePage.checklist)
        tlabel.place(x=30, y=275)
        ylabel = Label(self, text="%s" % ScorePage.checklist)
        ylabel.place(x=30, y=250)
        ulabel = Label(self, text="%s" % ScorePage.checklist)
        ulabel.place(x=30, y=225)
        ilabel = Label(self, text="%s" % ScorePage.checklist)
        ilabel.place(x=30, y=200)
        olabel = Label(self, text="%s" % ScorePage.checklist)
        olabel.place(x=30, y=175)
        plabel = Label(self, text="%s" % ScorePage.checklist)
        plabel.place(x=30, y=150)

        # print(ScorePage.checklist)

        savebutton =Button(self,text = 'Save', command = self.save)
        savebutton.place(x=30, y=420, width=170, height=50)

        bbutton = Button(self,text = 'Finish',command = self.quit)
        bbutton.place(x=215,y=420,width=170,height =50)

        abutton =Button(self,text = 'Restart',command = lambda :controller.show_frame("StartPage"))
        abutton.place(x=400, y=420, width=170, height=50)

        return {'info_label': userinfo, 'label_L': label, 'qlabel': qlabel, 'wlabel': wlabel, 'elabel': elabel, 'rlabel': rlabel, 'tlabel': tlabel, 'ylabel': ylabel, 'ulabel': ulabel, 'ilabel': ilabel, 'olabel': olabel, 'plabel': plabel, 'button_L': abutton, 'Bbutton':bbutton}

    # 오답 저장
    def save(self):
        # 파일명을 위해 User가 입력한 이름과 학번을 가져옴
        name_entry: Entry = StartPage.components['name_entry']
        user_name = name_entry.get()

        stid_entry: Entry = StartPage.components['stid_entry']
        student_id = stid_entry.get()

        # 이름과 학번을 입력하지 않았을 경우 파일명 변경
        if len(user_name)!=0 and len(student_id)!=0:
          filename = student_id+' '+user_name+' Checklist.xlsx'
        else:
          filename = 'Checklist.xlsx'

        workbook = xlsxwriter.Workbook(filename)

        # 현재 날짜와 시간을 바탕으로 시트명 설정
        sheetname = "%s%s%s_%s.%s.%s"%(datetime.now().year,datetime.now().month,datetime.now().day,datetime.now().hour,datetime.now().minute,datetime.now().second)
        worksheet = workbook.add_worksheet(sheetname)

        datalist = [["Word", "Answer"]]

        # datalist 데이터 추가
        for i in range(0,10):
          #from 0 to 9
          # print(i)
          if ScorePage.checklist[0][1][i][1]!=ScorePage.checklist[0][1][i][2]:
            datalist.append([ScorePage.checklist[0][1][i][0],ScorePage.checklist[0][1][i][1]])

        expenses = tuple(datalist)

        # 첫 번째 셀부터 값 입력
        row = 0
        col = 0

        # 엑셀 파일에 값 입력
        for word, answer in (expenses):
            worksheet.write(row, col, word)
            worksheet.write(row, col + 1, answer)
            row += 1

        workbook.close()

        #print("saved")

app = MyWindow()