from tkinter import *
import subprocess
import tkinter.messagebox
import datetime
import xlwt
from xlwt import Workbook


now = datetime.datetime.now()

def dashboard(g,pwdpresent):

    def back():
        root.destroy()

    def acc_details():
        for widget in mainframe.winfo_children():
            widget.destroy()
        Label(mainframe,text="Personal Details",font=("",16),fg="blue").grid(row=0,columnspan=2)
        Label(mainframe,text="Name :",font=("",12)).grid(row=1,column=0,sticky=N)
        Label(mainframe,text=name,font=("",12)).grid(row=1,column=1)
        Label(mainframe,text="Gender :",font=("",12)).grid(row=2,column=0,sticky=N)
        Label(mainframe,text=gender,font=("",12)).grid(row=2,column=1)
        Label(mainframe,text="Account Number :",font=("",12)).grid(row=3,column=0,sticky=N)
        Label(mainframe,text=acc_no,font=("",12)).grid(row=3,column=1)
        Label(mainframe,text="DOB :",font=("",12)).grid(row=4,column=0,sticky=N)
        Label(mainframe,text=dob,font=("",12)).grid(row=4,column=1)

        Label(mainframe,text="Contact Details",font=("",16),fg="blue").grid(row=5,columnspan=2)
        Label(mainframe,text="Address :",font=("",12)).grid(row=6,column=0,sticky=N)
        Label(mainframe,text=address,font=("",12)).grid(row=6,column=1)
        Label(mainframe,text="Email :",font=("",12)).grid(row=7,column=0,sticky=N)
        Label(mainframe,text=email,font=("",12)).grid(row=7,column=1)
        Label(mainframe,text="Phone Number :",font=("",12)).grid(row=8,column=0,sticky=N)
        Label(mainframe,text=phone,font=("",12)).grid(row=8,column=1)
        
    
    def balance_enquiry():
        for widget in mainframe.winfo_children():
            widget.destroy()

        p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
        p.communicate(input=b'y')

        readdet=open("det.txt","r").readlines()
        balance=readdet[0].split()

        p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
        p.communicate(input=b'y')
        
        bal="Current Balance is "+str(balance[int(g)-1])
        Label(mainframe,text=bal,font=("",14)).pack(side=LEFT,padx=30,pady=30)

    def transfer_money():

        def transfer():

            acc=account_entry.get()
            mon=money_entry.get()

            exitflag=0

            if acc.isdigit()==False:
                tkinter.messagebox.showinfo("Error","Enter valid Account Number")
                exitflag=1

            elif mon.isdigit()==False:
                tkinter.messagebox.showinfo("Error","Enter valid Amount")
                exitflag=1

                
            if exitflag==0:
                p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                readdet=open("det.txt","r").readlines()
                balance=readdet[0].split()
                lastno=len(balance)
                refno=int(int(acc)-350000)              

                if lastno<refno or refno<=0:
                    tkinter.messagebox.showinfo("Error","Invalid Account Number. Please recheck")

                elif refno==int(g):
                    tkinter.messagebox.showinfo("Error","Cannot Transfer to your Account. We didn't find that useful, so this feature is not added")

                elif int(mon) > int(balance[int(g)-1]):
                    tkinter.messagebox.showinfo("Error","Transfer Failed, Insufficient balance")

                else:
                    ded=(int(balance[int(g)-1]))-int(mon)
                    inc=(int(balance[refno-1]))+int(mon)
                    balance[int(g)-1]=str(ded)
                    balance[refno-1]=str(inc)
                    filewrite=open("det.txt","w")
                    for i in balance:
                        filewrite.write(i+" ")

                    filename=str(g)+".txt"
                    filename2=str(refno)+".txt"

                    p=subprocess.Popen('cacls '+filename+' /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                    
                    file=open(filename,"a")
                    file.write("Debit " +acc+" "+str(now.strftime("%d-%m-%Y@%H:%M:%S"))+" "+str(mon)+" "+str(ded)+"\n")

                    p=subprocess.Popen('cacls '+filename+' /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    p=subprocess.Popen('cacls '+filename2+' /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                    
                    file=open(filename2,"a")
                    file.write("Credit " +acc_no[:-1]+" "+str(now.strftime("%d-%m-%Y@%H:%M:%S"))+" "+str(mon)+" "+str(inc)+"\n")

                    p=subprocess.Popen('cacls '+filename2+' /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                        

                    tkinter.messagebox.showinfo("Success","Money Tranfer Successful")

                    transfer_money()

                p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
                p.communicate(input=b'y')



        for widget in mainframe.winfo_children():
            widget.destroy()

        Label(mainframe,text="Enter Account No.",font=("",11)).grid(row=0,column=0,padx=20,pady=20,sticky=E)
        account_entry=Entry(mainframe,bd=5)
        account_entry.grid(row=0,column=1,padx=10,pady=20)
        Label(mainframe,text="Enter Amount",font=("",11)).grid(row=1,column=0,padx=20,pady=20,sticky=E)
        money_entry=Entry(mainframe,bd=5)
        money_entry.grid(row=1,column=1,padx=10,pady=20)
        Button(mainframe,text="Submit",command=transfer).grid(columnspan=2,pady=20)


    def estat():
        def get_excel():
            s=Workbook()
            
            sheet=s.add_sheet('Sheet')
            p=subprocess.Popen('cacls '+file_ref+' /p everyone:f',stdin=subprocess.PIPE)
            p.communicate(input=b'y')
                        
            tranlist=open(file_ref,"r").readlines()
            length=len(tranlist)
            i=length-1
            exitflag=0
            if i==7:
                tkinter.messagebox.showinfo("Failed","No transactions have taken place")
                exitflag=1
            if exitflag==0:
                sheet.write(0,0,"DOMESTIC HOLDINGS")
                sheet.write(1,0,"Type")
                sheet.write(1,1,"Account No.")
                sheet.write(1,2,"Date and Time")
                sheet.write(1,3,"Amount")
                sheet.write(1,4,"Balance")
                j=2
                while i>7:
                    onetran=tranlist[i].split()
                    k=0
                    for element in onetran:
                        sheet.write(j,k,element)
                        k=k+1
                    j=j+1
                    i=i-1

            p=subprocess.Popen('cacls '+file_ref+' /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            name=g+".xls"

            s.save(name)

            tkinter.messagebox.showinfo("Success","Downloaded in Current Directory")
                
            
        for widget in mainframe.winfo_children():
            widget.destroy()

        p=subprocess.Popen('cacls '+file_ref+' /p everyone:f',stdin=subprocess.PIPE)
        p.communicate(input=b'y')
                    
        tranlist=open(file_ref,"r").readlines()
        length=len(tranlist)
        i=length-1
        exitflag=0
        if i==7:
            Label(mainframe,text="No Transaction yet. Better get going",font=("",13)).pack(side=LEFT,padx=20,pady=20)
            exitflag=1
        if exitflag==0:
            Label(mainframe,text="Type",font=("",10)).grid(row=0,column=0,padx=10,pady=10)
            Label(mainframe,text="Account No.",font=("",10)).grid(row=0,column=1,padx=10,pady=10)
            Label(mainframe,text="Date and Time",font=("",10)).grid(row=0,column=2,padx=10,pady=10)
            Label(mainframe,text="Amount",font=("",10)).grid(row=0,column=3,padx=10,pady=10)
            Label(mainframe,text="Balance",font=("",10)).grid(row=0,column=4,padx=10,pady=10)
            Button(mainframe,text='Export to excel',command=get_excel,bg="green",fg="white").grid(row=0,column=5,padx=20,pady=10)
            j=1
            while i>7:
                onetran=tranlist[i].split()
                k=0
                for element in onetran:
                    Label(mainframe,text=element,font=("",10)).grid(row=j,column=k,padx=10,pady=10)
                    k=k+1
                j=j+1
                i=i-1
            

        p=subprocess.Popen('cacls '+file_ref+' /p everyone:n',stdin=subprocess.PIPE)
        p.communicate(input=b'y')
        

    def change_pwd():
        for widget in mainframe.winfo_children():
            widget.destroy()

        def check_pwd():

            old_pwd=old_pwd_entry.get()
            new_pwd=new_pwd_entry.get()
            new_pwd2=new_pwd2_entry.get()

            if old_pwd!=pwdpresent:
                tkinter.messagebox.showinfo("Error","Enter correct current password")

            elif new_pwd!=new_pwd2:
                tkinter.messagebox.showinfo("Error","New and Confirm passwords mismatch")
            
            elif len(new_pwd)<8:
                tkinter.messagebox.showinfo("Password","Length of password must be atleast 8")

            else:
                p=subprocess.Popen('cacls login.txt /p everyone:f',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                fileread=open("login.txt","r")
                readlogin=fileread.readlines()
                entries=readlogin[int(g)-1].split()
                entries[1]=new_pwd
                readlogin[int(g)-1]=entries[0]+" "+entries[1]+" "+entries[2]+"\n"
                fileread.close()
                filewrite=open("login.txt","w")
                for i in readlogin:
                    filewrite.write(i)
                filewrite.close()

                p=subprocess.Popen('cacls login.txt /p everyone:n',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                tkinter.messagebox.showinfo("Sucess","Password change successful")
                change_pwd()


        Label(mainframe,text="Enter Current Password : ",font=("",10)).grid(row=0,column=0,padx=10,pady=10,sticky=E)
        old_pwd_entry=Entry(mainframe,bd=5,show='*')
        old_pwd_entry.grid(row=0,column=1,padx=10,pady=10)
        Label(mainframe,text="Enter New Password : ",font=("",10)).grid(row=1,column=0,padx=10,pady=10,sticky=E)
        new_pwd_entry=Entry(mainframe,bd=5,show='*')
        new_pwd_entry.grid(row=1,column=1,padx=10,pady=10)
        Label(mainframe,text="Confirm New Password : ",font=("",10)).grid(row=2,column=0,padx=10,pady=10,sticky=E)
        new_pwd2_entry=Entry(mainframe,bd=5,show='*')
        new_pwd2_entry.grid(row=2,column=1,padx=10,pady=10)
        Button(mainframe,text="Submit",command=check_pwd).grid(columnspan=2,padx=10,pady=20)


    def view_invest():
        
        for widget in mainframe.winfo_children():
            widget.destroy()

       
        p=subprocess.Popen('cacls investments.txt /p everyone:f',stdin=subprocess.PIPE)
        p.communicate(input=b'y')
                    
        list1=open("investments.txt","r").readlines()

        p=subprocess.Popen('cacls investments.txt /p everyone:n',stdin=subprocess.PIPE)
        p.communicate(input=b'y')

        acctest=str(350000+int(g))

        count=0

        for i in list1:
            q=i.split()
            if q[0]==acctest:
                count=count+1

        if count>0:

            p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            comlist=open("sharehold.txt","r").readlines()

            p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')
                    
            Label(mainframe,text="S.No",font=("",12)).grid(row=0,column=0,padx=10,pady=10)
            Label(mainframe,text="Comapny ID",font=("",12)).grid(row=0,column=1,padx=10,pady=10)
            Label(mainframe,text="No. of Shares",font=("",12)).grid(row=0,column=2,padx=10,pady=10)
            Label(mainframe,text="Bought value",font=("",12)).grid(row=0,column=3,padx=10,pady=10)
            Label(mainframe,text="Present Share",font=("",12)).grid(row=0,column=4,padx=10,pady=10)
            Label(mainframe,text="Value",font=("",12)).grid(row=0,column=5,padx=10,pady=10)

            j=1
            acctest=str(350000+int(g))
            

            for i in list1:
                entry=i.split()
                if str(entry[0])==acctest:
                    Label(mainframe,text=j,font=("",10)).grid(row=j,column=0,padx=10,pady=10)
                    Label(mainframe,text=entry[1],font=("",10)).grid(row=j,column=1,padx=10,pady=10)
                    Label(mainframe,text=entry[2],font=("",10)).grid(row=j,column=2,padx=10,pady=10)
                    Label(mainframe,text=entry[3],font=("",10)).grid(row=j,column=3,padx=10,pady=10)
                    present_value=0.0
                    for q in comlist:
                        name=q.split()
                        if entry[1]==name[1]:
                            present_value=float(name[3])
                            break
                    Label(mainframe,text=str(present_value),font=("",10)).grid(row=j,column=4,padx=10,pady=10)
                    value=int(present_value*int(float(entry[2])))
                    Label(mainframe,text=str(value),font=("",10)).grid(row=j,column=5,padx=10,pady=10)
                    j=j+1

            Label(mainframe,text="Enter company ID to sell shares",font=("",12)).grid(row=j+1,columnspan=3,padx=10,pady=20,sticky=E)

            id_entry=Entry(mainframe,bd=5)
            id_entry.grid(row=j+1,column=3,columnspan=2,padx=10,pady=20)

            Label(mainframe,text="Enter no. of shares to sell",font=("",12)).grid(row=j+2,columnspan=3,padx=10,pady=20,sticky=E)

            share_entry=Entry(mainframe,bd=5)
            share_entry.grid(row=j+2,column=3,columnspan=2,padx=10,pady=20)

            def sell_stocks():

                comid=id_entry.get()
                share=share_entry.get()
                shareno=int(share)

                p=subprocess.Popen('cacls investments.txt /p everyone:f',stdin=subprocess.PIPE)
                p.communicate(input=b'y')
                    
                list1=open("investments.txt","r").readlines()

                p=subprocess.Popen('cacls investments.txt /p everyone:n',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                acctest=str(350000+int(g))

                m=0
                flag=0

                for i in list1:
                    entry=i.split()
                    if str(entry[0])==acctest and str(entry[1])==comid and int(str(entry[2]))>=shareno:
                        if int(str(entry[2]))==shareno:
                            p=subprocess.Popen('cacls investments.txt /p everyone:f',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                            p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                            lines=open("sharehold.txt","r").readlines()
                    
                            f=open("investments.txt","w")

                            pres_value=0.0
                            
                            for k in list1:
                                if k==list1[m]:
                                    h=lines[m].split()
                                    pres_value=float(h[3])
                                    continue
                                f.write(k)

                            f.close()

                            p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                            p=subprocess.Popen('cacls investments.txt /p everyone:n',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                        else:
                            p=subprocess.Popen('cacls investments.txt /p everyone:f',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                            p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                            lines=open("sharehold.txt","r").readlines()
                    
                            f=open("investments.txt","w")

                            for k in list1:
                                if k==list1[m]:
                                    dummy=k.split()
                                    dummy[2]=str(int(dummy[2])-shareno)
                                    z=""
                                    h=lines[m].split()
                                    pres_value=float(h[3])
                                    for e in dummy:
                                        z+=e+" "
                                    f.write(z+'\n')
                                    continue
                                f.write(k)

                            f.close()

                            p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')

                            p=subprocess.Popen('cacls investments.txt /p everyone:n',stdin=subprocess.PIPE)
                            p.communicate(input=b'y')
                            


                        p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
                        p.communicate(input=b'y')
                    
                        ballist=open("det.txt","r").readlines()
                        
                        bal=ballist[0].split()

                        deposit_value=int(shareno*pres_value)

                        r=str(int(bal[int(g)-1])+deposit_value)

                        bal[int(g)-1]=r

                        f=open("det.txt","w")

                        for i in bal:
                            f.write(i+" ")

                        f.close()

                        p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
                        p.communicate(input=b'y')


                        p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
                        p.communicate(input=b'y')

                        lines=open("sharehold.txt","r").readlines()

                        z=""
                        for i in lines:
                            l=i.split()
                            if l[1]==comid:
                                l[2]=str(int(l[2])+shareno)
                                for n in l:
                                    z+=str(n)+" "
                                break

                        f=open("sharehold.txt","w")
                        
                        for i in lines:
                            zs=z.split()
                            iss=i.split()
                            if zs[0]==iss[0]:
                                f.write(z+"\n")
                                continue
                            f.write(i)

                        f.close()

                        p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
                        p.communicate(input=b'y')

                        filename=g+".txt"
                        
                        p=subprocess.Popen('cacls '+filename+' /p everyone:f',stdin=subprocess.PIPE)
                        p.communicate(input=b'y')

                        f=open(filename,"a")

                        now = datetime.datetime.now()

                        f.write("Credit Investment "+str(now.strftime("%d-%m-%Y@%H:%M:%S"))+" "+str(deposit_value)+" "+str(int(r))+"\n")

                        f.close()

                        p=subprocess.Popen('cacls '+filename+' /p everyone:n',stdin=subprocess.PIPE)
                        p.communicate(input=b'y')

                        tkinter.messagebox.showinfo("Success","Money transferred to your bank")

                        flag=1

                        view_invest()

                        break


                    m=m+1
                    
                if flag==0:

                    tkinter.messagebox.showinfo("Error","Please recheck values")


            Button(mainframe,text="Sell stocks",command=sell_stocks).grid(row=j+3,columnspan=6,padx=20,pady=40)

        else:
            Label(mainframe,text="No Investments at present. Go to Invest option to invest in a company",font=("",14)).grid(row=0,column=0,padx=20,pady=20)
                


    def invest():

        def check_invest():

            def investment():
                
                exit_flag1=0
                noshares=getshares.get()
                amount=int(noshares)
                sharesleft=int(entry[2])
                sharevalue=float(entry[3])

                dedamt=float(sharevalue*amount*1.05)

                p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                detopen=open("det.txt","r")
                bal_lines=detopen.readlines()
                bal=bal_lines[0].split()
                detopen.close()

                if amount>int(entry[2]):
                    tkinter.messagebox.showinfo("Notice","Max. shares limit exceeded. The left no. of shares are "+entry[2]+" ")
                    exit_flag1=1
                
                elif float(bal[int(g)-1])<dedamt:
                    tkinter.messagebox.showinfo("Error","Insufficient balance")
                    exit_flag1=1

                if exit_flag1==0:
                    i=0
                    p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    comlist=open("sharehold.txt","r").readlines()

                    p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                    

                    for u in comlist:
                        data=u.split()
                        if data[1]==getid:
                            entry1=data
                            break
                        i=i+1

                    entry1[2]=str(int(entry1[2])-amount)
                    string=""
                    for j in entry1:
                        string+=str(j)+" "

                    comlist[int(i)]=string

                    p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    f=open("sharehold.txt","w")

                    for k in comlist:
                        dummy=k.split()
                        for i in dummy:
                            f.write(i+' ')
                        f.write('\n')
                    f.close()

                    p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                    
                    x=str(int(int(bal[int(g)-1])-int(dedamt)))

                    bal[int(g)-1]=x

                    f=open("det.txt","w")
                    
                    for i in bal:
                        f.write(i+" ")
                    f.close()

                    filename=str(g)+".txt"

                    p=subprocess.Popen('cacls '+filename+' /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                    
                    file=open(filename,"a")
                    file.write("Debit Investment "+str(now.strftime("%d-%m-%Y@%H:%M:%S"))+" "+str(int(dedamt))+" "+str(x)+"\n")
                    file.close()

                    p=subprocess.Popen('cacls '+filename+' /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    p=subprocess.Popen('cacls investments.txt /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    f=open("investments.txt","a")

                    f.write(str(350000+int(g))+" "+getid+" "+str(amount)+" "+str(sharevalue)+"\n")

                    f.close()

                    p=subprocess.Popen('cacls investments.txt /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    tkinter.messagebox.showinfo("Success",str(amount)+" no .of shares have been invested with 5% stock brokerage summing up to "+str(int(dedamt)))


                p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                invest()

                    
            getid=getdata.get()
            exit_flag=0
            
            if getid.isalpha()==False or len(getid)!=3:
                  tkinter.messagebox.showinfo("Error","Enter valid ID (Alphabetic of length 3)")
                  exit_flag=1
                  
            entry=[]

            for u in comlist:
                data=u.split()
                if data[1]==getid:
                    entry=data
                    break

            if len(entry)==0:
                tkinter.messagebox.showinfo("Error","No matching company detected")
                exit_flag=1
            
            if exit_flag==0:
                for widget in mainframe.winfo_children():
                    widget.destroy()
                Label(mainframe,text="Name of Company",font=("",12)).grid(row=0,column=0,padx=10,pady=10)
                Label(mainframe,text=entry[0],font=("",12)).grid(row=0,column=1,padx=10,pady=10)
                Label(mainframe,text="Company ID",font=("",12)).grid(row=1,column=0,padx=10,pady=10)
                Label(mainframe,text=entry[1],font=("",12)).grid(row=1,column=1,padx=10,pady=10)
                Label(mainframe,text="Shares Left",font=("",12)).grid(row=2,column=0,padx=10,pady=10)
                Label(mainframe,text=entry[2],font=("",12)).grid(row=2,column=1,padx=10,pady=10)
                Label(mainframe,text="Present Share",font=("",12)).grid(row=3,column=0,padx=10,pady=10)
                Label(mainframe,text=entry[3],font=("",12)).grid(row=3,column=1,padx=10,pady=10)

                Label(mainframe,text="Enter no. of shares to be invested",font=("",12)).grid(row=4,columnspan=2,padx=10,pady=20)
                getshares=Entry(mainframe,bd=5)
                getshares.grid(row=4,column=2,padx=10,pady=20)
                Button(mainframe,text="Invest",command=investment).grid(row=5,columnspan=2,padx=10,pady=10)

        for widget in mainframe.winfo_children():
            widget.destroy()
        

        p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
        p.communicate(input=b'y')

        comlist=open("sharehold.txt","r").readlines()

        p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
        p.communicate(input=b'y')

        if(len(comlist))>0:

            Label(mainframe,text="S.No",font=("",12)).grid(row=0,column=0,padx=10,pady=10)
            Label(mainframe,text="Name of the Company",font=("",12)).grid(row=0,column=1,padx=10,pady=10)
            Label(mainframe,text="Company ID",font=("",12)).grid(row=0,column=2,padx=10,pady=10)
            Label(mainframe,text="No. of Shares Left",font=("",12)).grid(row=0,column=3,padx=10,pady=10)
            Label(mainframe,text="Present Share Hold",font=("",12)).grid(row=0,column=4,padx=10,pady=10)

            j=1
            for u in comlist:
                entry=u.split()
                Label(mainframe,text=j,font=("",10)).grid(row=j,column=0,padx=10,pady=10)
                Label(mainframe,text=entry[0],font=("",10)).grid(row=j,column=1,padx=10,pady=10)
                Label(mainframe,text=entry[1],font=("",10)).grid(row=j,column=2,padx=10,pady=10)
                Label(mainframe,text=entry[2],font=("",10)).grid(row=j,column=3,padx=10,pady=10)
                Label(mainframe,text=entry[3],font=("",10)).grid(row=j,column=4,padx=10,pady=10)

                j=j+1

            Label(mainframe,text="Enter Company ID to invest",font=("",12)).grid(row=j+1,columnspan=2,padx=10,pady=20,sticky=E)
            getdata=Entry(mainframe,bd=5)
            getdata.grid(row=j+1,column=2,padx=10,pady=20)
            Button(mainframe,text="Invest",font=("",10),command=check_invest).grid(row=j+2,columnspan=4,padx=10,pady=10)

        else:
            Label(mainframe,text="No companies are listed for stock exchange",font=("",14)).grid(row=0,column=0,padx=20,pady=20)

        

    file_ref=str(g)+".txt"

    #unlock
    p=subprocess.Popen('cacls '+file_ref+' /p everyone:f',stdin=subprocess.PIPE)
    p.communicate(input=b'y')

    file=open(file_ref,"r")
    details=file.readlines()
    acc_no=details[0]
    name=details[1]
    dob=details[2]
    gender=details[3]
    address=details[4]
    email=details[5]
    phone=details[6]
    aadhar=details[7]
    

    #lock
    p=subprocess.Popen('cacls '+file_ref+' /p everyone:n',stdin=subprocess.PIPE)
    p.communicate(input=b'y')

    root = Tk()

    root.overrideredirect(True)
    root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))

    root.title('Dashboard')

    top_frame = Frame(root, bg="black")

    bank_name = Label(top_frame, text="Domestic Holdings Inc.", bg="black",fg="white", font=("",13))
    bank_name.pack(side=LEFT)
    quitButton = Button(top_frame, text='X Close', bg="red", command=quit)
    quitButton.pack(side=RIGHT)
    logoutButton = Button(top_frame, text='-> Log Out', bg="black",fg="white", command=login)
    logoutButton.pack(side=RIGHT)

    top_frame.pack(fill=X)

    det_frame=Frame(root,height=200,bg="black")
    det_frame.pack(fill=X,side=TOP)

    det_name_label=Label(det_frame,text="\n"+"Welcome "+name+"\n",bg="black",fg="white",font=("times new roman",15))
    det_name_label.pack(side=RIGHT)

    func_frame=Frame(root,width=200,bg="black")
    func_frame.pack(fill=Y,side=LEFT)

    mainframe=Frame(root)
    mainframe.pack(side=TOP,fill=BOTH)

    balance_enquiry_button=Button(func_frame,text="Balance Enquiry",width=25,command=balance_enquiry,bg="black",fg="white")
    balance_enquiry_button.pack(pady=10)

    acc_details_button=Button(func_frame,text="Account Details",width=25,command=acc_details,bg="black",fg="white")
    acc_details_button.pack(pady=10)

    withdraw_button = Button(func_frame, text="e-Statement", width=25,command=estat,bg="black",fg="white")
    withdraw_button.pack(pady=10)

    transfer_button = Button(func_frame, text="Transfer Money", width=25,command=transfer_money,bg="black",fg="white")
    transfer_button.pack(pady=10)

    change_pwd_button = Button(func_frame, text="Change Password", width=25,command=change_pwd,bg="black",fg="white")
    change_pwd_button.pack(pady=10)

    Button(func_frame, text="View Investments", width=25,command=view_invest,bg="black",fg="white").pack(pady=10)

    Button(func_frame, text="Invest", width=25,command=invest,bg="black",fg="white").pack(pady=10)

    root.mainloop()
    

def manager_dashboard():

    def depo():

        def deposit():
            acc=acc_entry.get()
            mon=amount_entry.get()

            exitflag=0

            if acc.isdigit()==False:
                tkinter.messagebox.showinfo("Error","Enter valid Account Number")
                exitflag=1

            elif mon.isdigit()==False:
                tkinter.messagebox.showinfo("Error","Enter valid Amount")
                exitflag=1

                
            if exitflag==0:
                p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
                p.communicate(input=b'y')

                readdet=open("det.txt","r").readlines()
                balance=readdet[0].split()
                lastno=len(balance)
                refno=int(int(acc)-350000)
                filename=str(refno)+".txt"

                if lastno<refno or refno<=0:
                    tkinter.messagebox.showinfo("Error","Invalid Account Number. Please recheck")


                else:
                    inc=(int(balance[refno-1]))+int(mon)
                    balance[refno-1]=str(inc)
                    filewrite=open("det.txt","w")
                    for i in balance:
                        filewrite.write(i+" ")

                    p=subprocess.Popen('cacls '+filename+' /p everyone:f',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')
                        
                    file=open(filename,"a")
                    file.write("Credit " +"Self"+" "+str(now.strftime("%d-%m-%Y@%H:%M:%S"))+" "+str(mon)+" "+str(inc)+"\n")

                    p=subprocess.Popen('cacls '+filename+' /p everyone:n',stdin=subprocess.PIPE)
                    p.communicate(input=b'y')

                    tkinter.messagebox.showinfo("Success","Money Deposit Successful")

                    depo()

        for widget in bottom_frame.winfo_children():
            widget.destroy()

        p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
        p.communicate(input=b'y')

        acc_name=Label(bottom_frame,text="Account Number")
        acc_name.grid(row=0,column=0,padx=0,pady=20)
        acc_entry=Entry(bottom_frame,bd=5,width=20)
        acc_entry.grid(row=0,column=1,padx=20)

        amount_name=Label(bottom_frame,text="Deposit Amount")
        amount_name.grid(row=1,column=0,pady=10)
        amount_entry=Entry(bottom_frame,bd=5,width=20)
        amount_entry.grid(row=1,column=1,padx=20,pady=10)

        butn=Button(bottom_frame,text="Deposit",command=deposit)
        butn.grid(columnspan=2,pady=20,padx=20)


    def change_share():

        def changevalue():

            id_get=id_entry.get()
            share_get=share_entry.get()
            
            p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            sharelist=open("sharehold.txt","r").readlines()


            z=""
            for u in sharelist:
                k=u.split()
                if k[1]==id_get:
                    k[3]=share_get
                    for x in k:
                        z+=x+" "

            f=open("sharehold.txt","w")

            for u in sharelist:
                k=u.split()
                if k[1]==id_get:
                    f.write(z+"\n")
                    continue
                f.write(u)

            f.close()

            tkinter.messagebox.showinfo("Success","Changed share price successfully")

            id_entry.delete(0,END)
            share_entry.delete(0,END)

            p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')
                    
            
        for widget in bottom_frame.winfo_children():
            widget.destroy()


        Label(bottom_frame,text="Enter company ID",font=("",10)).grid(row=0,column=0,padx=10,pady=10,sticky=E)

        id_entry=Entry(bottom_frame,bd=5)
        id_entry.grid(row=0,column=1,padx=10,pady=10)

        Label(bottom_frame,text="Enter share value",font=("",10)).grid(row=1,column=0,padx=10,pady=10,sticky=E)

        share_entry=Entry(bottom_frame,bd=5)
        share_entry.grid(row=1,column=1,padx=10,pady=10)

        Button(bottom_frame,text="Submit",command=changevalue).grid(row=2,columnspan=2,padx=10,pady=20)


    def add_company():

        def add_c():

            name_get=name_entry.get()
            id_get=id_entry.get()
            maxshares_get=maxshares_entry.get()
            shareval_get=shareval_entry.get()

            p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            f=open("sharehold.txt","a")

            f.write(name_get+" "+id_get+" "+maxshares_get+" "+shareval_get+"\n")

            f.close()

            tkinter.messagebox.showinfo("Success","Comapny added successfully")

            add_company()

        for widget in bottom_frame.winfo_children():
            widget.destroy()

        Label(bottom_frame,text="Enter Company Name",font=("",10)).grid(row=0,column=0,padx=10,pady=10)

        name_entry=Entry(bottom_frame,bd=5)
        name_entry.grid(row=0,column=1,padx=10,pady=10)

        Label(bottom_frame,text="Enter Company ID",font=("",10)).grid(row=1,column=0,padx=10,pady=10)

        id_entry=Entry(bottom_frame,bd=5)
        id_entry.grid(row=1,column=1,padx=10,pady=10)

        Label(bottom_frame,text="Enter Maximum shares",font=("",10)).grid(row=2,column=0,padx=10,pady=10)

        maxshares_entry=Entry(bottom_frame,bd=5)
        maxshares_entry.grid(row=2,column=1,padx=10,pady=10)

        Label(bottom_frame,text="Enter Share Value",font=("",10)).grid(row=3,column=0,padx=10,pady=10)

        shareval_entry=Entry(bottom_frame,bd=5)
        shareval_entry.grid(row=3,column=1,padx=10,pady=10)

        Button(bottom_frame,text="Add",command=add_c).grid(row=4,columnspan=2,padx=10,pady=20)
        

            
                
    root =Tk()
    root.overrideredirect(True)
    root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))
    root.title('dashboard')

    frame=Frame(root,bg="black")
    bank_name=Label(frame,text="Domestic Holdings Inc.",bg="black",fg="white")
    bank_name.pack(side=LEFT)
    Button(frame,text='X Close',bg="red",command=quit).pack(side=RIGHT)
    Button(frame,text="->] Log Out",fg="white",bg="black",command=login).pack(side=RIGHT)
    frame.pack(fill=BOTH)

    top_frame=Frame(root,bg="black",pady=30)
    top_frame.pack(fill=X,side=TOP)
    top_frame_det=Label(top_frame,text="Welcome Manager,Domestic Holdings Inc.",bg="black",fg="white",font=("",20))
    top_frame_det.pack()

    func_frame=Frame(root,bg="black")
    func_frame.pack(fill=Y,side=LEFT)

    bottom_frame=Frame(root)
    bottom_frame.pack(fill=BOTH)

    Button(func_frame,text="Deposits",font=("",10),command=depo,width=25,fg="white",bg="black").pack(padx=10,pady=10)

    Button(func_frame,text="Change Share Value",font=("",10),command=change_share,width=25,fg="white",bg="black").pack(padx=10,pady=10)

    Button(func_frame,text="Add company",font=("",10),command=add_company,width=25,fg="white",bg="black").pack(padx=10,pady=10)


    root.mainloop()



def register():

    def get_reg_details():

        def generate_userid(ref):
            user=str(350000+flag)
            return user

        def generate_password(ref):
            pwd=name[:-4:-1]
            pwd+=aadhar[:-5:-1]
            pwd+=str(ref)
            return pwd
            
            
        name = name_entry.get()
        dob = dob_entry.get()
        address = address_entry.get()
        email = email_entry.get()
        phone = phone_entry.get()
        aadhar = aadhar_entry.get()
        gender=gender_select.get()

        exit_flag = 0
        flag = 0
        if name == "" or dob == "" or address == "" or email == "" or phone == "" or aadhar == "":
            exit_flag = 1
            tkinter.messagebox.showinfo('Error','All fields are mandatory. Make sure all are filled')

        elif len(name)<3:
            exit_flag=1
            tkinter.messagebox.showinfo('Error','Enter valid name')
            

        elif dob.isdigit()==False or len(dob)!=8:
            exit_flag=1
            tkinter.messagebox.showinfo('Error','DOB format is incorrect')

        elif int(dob[:2])>31 or int(dob[2:-4])>12 or int(dob[4:])>2000:
            exit_flag=1
            tkinter.messagebox.showinfo('Error','Date in DOB is incorrect or Insufficient age for opening of account')

        elif gender=="------":
            exit_flag=1
            tkinter.messagebox.showinfo('Error','Select Gender')

        elif email.find('@')==-1 or email.find('.com')==-1 or email.find('.com')<email.find('@'):
            exit_flag=1
            tkinter.messagebox.showinfo('Error','Enter valid email')

        elif len(phone)!=13 or phone[1:].isdigit()==False:
            exit_flag=1
            tkinter.messagebox.showinfo('Error','Enter valid Phone no.')

        elif len(aadhar)!=12 or aadhar.isdigit() == False:
            exit_flag=1
            tkinter.messagebox.showinfo('Error','Enter valid Aadhar no.')
        

        if exit_flag == 0:
            
            #unlock
            p=subprocess.Popen('cacls login.txt /p everyone:f',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            flag=len(open("login.txt","r").readlines())
            flag = flag + 1
            userid = generate_userid(flag)
            password = generate_password(flag)
            file = open("login.txt", "a")
            file.write(str(userid) + " " + str(password) + " " + str(flag) + "\n")
            file.close()

            #lock
            p=subprocess.Popen('cacls login.txt /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')
            
            filename = str(flag) + ".txt"
            file = open(filename, "w")
            file.write(str(350000+flag)+ "\n")
            file.write(name + "\n" + dob + "\n" + gender + "\n" + address + "\n" + email + "\n" + phone + "\n"+aadhar+"\n")
            file.close()

            #lock user
            p=subprocess.Popen('cacls '+str(flag)+'.txt /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            #unlock det
            p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
            p.communicate(input=b'y')

            file=open("det.txt","a")
            file.write("1000 ")
            file.close()

            #lock det
            p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')
            
            tkinter.messagebox.showinfo('Success', 'Registered Successfully\nYour UserID is '+userid+'\nPassword is '+password+'\nRemember to login')
            tkinter.messagebox.showwarning('**Important**', 'Remember your UsenID and password\nYour UserIDis '+userid+'\nPassword is '+password)
            goback()


    def goback():
        root.destroy()
        login()


    root = Tk()


    root.overrideredirect(True)
    root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))

    root.title('Registration')

    top_frame = Frame(root, bg="black")

    bank_name = Label(top_frame, text="Domestic Holdings Inc.", bg="black",fg="white",font=("",13))
    bank_name.pack(side=LEFT)
    quitButton = Button(top_frame, text='X Close', bg="red", command=quit)
    quitButton.pack(side=RIGHT)

    top_frame.pack(fill=BOTH)

    RDframe1 = Frame(root, height=100)
    Rframe1 = Frame(root)

    back_button = Button(top_frame, text="< Back to Login Page", fg="white",bg="black", command=goback)
    back_button.pack(side=RIGHT)

    regtitle = Label(Rframe1, text="REGISTRATION FORM")
    regtitle.config(font=("", 25))
    regtitle.pack(side=TOP)

    Rframe2 = Frame(root, padx=50, pady=50)

    name_label = Label(Rframe2, text="Name :")
    name_label.grid(row=0, column=0, padx=10, pady=10, sticky=E)

    name_entry = Entry(Rframe2, bd=5,width=30)
    name_entry.grid(row=0, column=1, pady=10, padx=10, sticky=W)

    dob_label = Label(Rframe2, text="Date of Birth (DDMMYYYY) :")
    dob_label.grid(row=1, column=0, pady=10, padx=10, sticky=E)

    dob_entry = Entry(Rframe2, bd=5,width=30)
    dob_entry.grid(row=1, column=1, padx=10, pady=10, sticky=W)

    gender_label = Label(Rframe2, text="Gender :")
    gender_label.grid(row=2, column=0, pady=10, padx=10, sticky=E)

    gender_select = StringVar(Rframe2)  # drop down menu
    gender_select.set("------")  # default value
    w = OptionMenu(Rframe2, gender_select, "Male", "Female")
    w.grid(row=2, column=1, padx=10, pady=10, sticky=W)

    address_label = Label(Rframe2, text="Address :")
    address_label.grid(row=3, column=0, padx=10, pady=5, sticky=E)

    address_entry = Entry(Rframe2, bd=5,width=30)
    address_entry.grid(row=3, column=1, pady=10, padx=10, sticky=W)

    email_label = Label(Rframe2, text="E-Mail :")
    email_label.grid(row=4, column=0, pady=10, padx=10, sticky=E)

    email_entry = Entry(Rframe2, bd=5,width=30)
    email_entry.grid(row=4, column=1, pady=10, padx=10, sticky=W)

    phone_label = Label(Rframe2, text="Phone No. :")
    phone_label.grid(row=5, column=0, pady=10, padx=10, sticky=E)

    phone_entry = Entry(Rframe2, bd=5,width=30)
    phone_entry.insert(0, "+91")
    phone_entry.grid(row=5, column=1, pady=10, padx=10, sticky=W)

    aadhar_label = Label(Rframe2, text="Aadhar No. :")
    aadhar_label.grid(row=6, column=0, pady=10, padx=10, sticky=E)

    aadhar_entry = Entry(Rframe2, bd=5,width=30)
    aadhar_entry.grid(row=6, column=1, pady=10, padx=10, sticky=W)

    submit_button = Button(Rframe2, text="Submit", command=get_reg_details)
    submit_button.grid(columnspan=2, pady=10)

    reg_note = Label(Rframe2, text="*All fields are mandatory", fg="red")
    reg_note.config(font=("", 12))
    reg_note.grid(columnspan=2, pady=30)

    RDframe1.pack(side=TOP)
    Rframe1.pack(side=TOP)
    Rframe2.pack(side=TOP)

    root.mainloop()



def login():
    g=0

    def get_pwd():
        a=user_name_entry.get()
        b=password_entry.get()
        flag=0

        if a=='master' and b=='master':
            root.destroy()
            manager_dashboard()

        #unlock
        p=subprocess.Popen('cacls login.txt /p everyone:f',stdin=subprocess.PIPE)
        p.communicate(input=b'y')
        

        for line in open("login.txt", "r").readlines():
            login_info = line.split()
            if a == login_info[0] and b == login_info[1]:
                #lock
                p=subprocess.Popen('cacls login.txt /p everyone:n',stdin=subprocess.PIPE)
                p.communicate(input=b'y')
                flag=1
                user_name_entry.delete(0,END)
                password_entry.delete(0,END)
                root.destroy()
                dashboard(login_info[2],login_info[1])
        if flag==0:
            tkinter.messagebox.showinfo('Error', 'Invalid Details')
            #lock
            p=subprocess.Popen('cacls login.txt /p everyone:n',stdin=subprocess.PIPE)
            p.communicate(input=b'y')



    def clear():
        root.destroy()
        register()


    root = Tk()

    root.overrideredirect(True)
    root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))

    root.title('Login')

    top_frame=Frame(root,bg="black")

    bank_name = Label(top_frame,text="Domestic Holdings Inc.",bg="black",fg="white",font=("",13))
    bank_name.pack(side=LEFT)
    quitButton=Button(top_frame,text='X Close',bg="red",command=quit)
    quitButton.pack(side=RIGHT)

    top_frame.pack(fill=BOTH)

    LDframe1=Frame(root,height=120)
    LDframe1.pack(fill=X)

    Lframe1=Frame(root)
    Lframe1.pack(side=TOP,padx=100,pady=50)

    login_label=Label(Lframe1,text="LOGIN")
    login_label.config(font=("",25))
    login_label.grid(columnspan=2,pady=30)

    user_name_label = Label(Lframe1,text = 'UserID:')
    user_name_label.grid(row=1,column=0,padx=5,pady=20)

    user_name_entry = Entry(Lframe1,bd =5)
    user_name_entry.grid(row=1,column=1,padx=5,pady=20)

    password_label = Label(Lframe1,text = 'Password:')
    password_label.grid(row=2,column=0,padx=5,pady=20)

    password_entry = Entry(Lframe1, bd=5, show='*')
    password_entry.grid(row=2,column=1,padx=5,pady=20)

    login_button = Button(Lframe1, text = 'Login',command = get_pwd)
    login_button.grid(row=3,column=1,pady=20)

    reg_button=Button(Lframe1,text="New User? Register",command=clear)
    reg_button.grid(row=4,column=1)

    LDframe2=Frame(root,height=30,padx=100,pady=50)
    LDframe2.pack(side=BOTTOM)

    Lframe2 = Frame(root,padx=100,pady=50)
    Lframe2.pack(side=BOTTOM, fill=X)

    notes = Label(Lframe2,text="Enter your username and password to avail banking services")
    caution=Label(Lframe2,text="Never share your username and password. Doing so, you may lose the aceess of your account",fg="red")

    caution.pack(side=BOTTOM)
    notes.pack(side=BOTTOM)

    root.mainloop()


p=subprocess.Popen('cacls login.txt /p everyone:f',stdin=subprocess.PIPE)
p.communicate(input=b'y')
file=open("login.txt","a")
file.write("")
file.close()
p=subprocess.Popen('cacls login.txt /p everyone:n',stdin=subprocess.PIPE)
p.communicate(input=b'y')

p=subprocess.Popen('cacls investments.txt /p everyone:f',stdin=subprocess.PIPE)
p.communicate(input=b'y')
file=open("investments.txt","a")
file.write("")
file.close()
p=subprocess.Popen('cacls investments.txt /p everyone:n',stdin=subprocess.PIPE)
p.communicate(input=b'y')

p=subprocess.Popen('cacls sharehold.txt /p everyone:f',stdin=subprocess.PIPE)
p.communicate(input=b'y')
file=open("sharehold.txt","a")
file.write("")
file.close()
p=subprocess.Popen('cacls sharehold.txt /p everyone:n',stdin=subprocess.PIPE)
p.communicate(input=b'y')

p=subprocess.Popen('cacls det.txt /p everyone:f',stdin=subprocess.PIPE)
p.communicate(input=b'y')
file=open("det.txt","a")
file.write("")
file.close()
p=subprocess.Popen('cacls det.txt /p everyone:n',stdin=subprocess.PIPE)
p.communicate(input=b'y')

login()
