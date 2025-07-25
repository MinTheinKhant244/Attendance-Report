from docxtpl import DocxTemplate

def save(lists):
    doc=DocxTemplate('students_attendance1.docx')
    doc.render({
    "slists":lists})
    doc.save('new_document.docx')
    print("Saved Successfully.")

def show(lists):
    if lists is None or len(lists) == 0:
        print("Empty...")
    else :
        for li in lists:
            print(li,end="\n")

def delete(lists,ind):
    if (ind<=len(lists)):
        del(lists[ind-1])
        print()
    else:
        print("Your index doesn't exist")
    

def edit(lists,ind):
    index=ind-1
    count=ind
    name=(input("Student's name: "))
    start_date="June 1, 2024"
    while True:
        try:
            att=int(input("Enter attendance (0-60):"))
            if att<0:
                print("Please enter carefully!!!")
                continue
            elif att>60:
                print("Please enter carefully!!!")
                continue
            else:
                break 
        except:
            print("Please enter carefully!!!")
            continue
    all=60
    attendance=f"{att}/{all}"
    per=(att/all)*100
    percen=round(per)
    percentage=f"{percen}%"          
    remark=" "
    add=[count,name,start_date,attendance,percentage,remark]
    lists[index]=add
    print("Edited Successfully")
    print(lists[index])
    
def create(lists,count):
    name=input("Student's name: ")
    start_date="June 1, 2024"
    while True:
        try:
            att=int(input("Enter attendance (0-60):"))
            if att<0:
                print("Please enter carefully!!!")
                continue
            elif att>60:
                print("Please enter carefully!!!")
                continue
            else:
                break 
        except:
            print("Please enter carefully!!!")
            continue   
    all=60
    attendance=f"{att}/{all}"
    per=(att/all)*100
    percen=round(per)
    percentage=f"{percen}%"          
    remark=" "
    add=[count,name,start_date,attendance,percentage,remark]
    lists.append(add)
    print(add)

def main():
    lists=[]
    count=1
    while True:
        print("1. create list")
        print("2. edit list")
        print("3. View all lists")
        print("4. Delete")
        print("5. Delete All")
        print("6. Save to Docx & exit")
        choice=input("Enter Your choice(1-6): ")
        if choice=='1':
            create(lists,count)
            count+=1
            
        elif choice=='2':
            while True:
                try:
                    ind=int(input("Enter place you want to edit: "))
                    if ind<=len(lists):
                        edit(lists,ind)
                        break
                    else:
                        print("Empty.")
                        break
                except ValueError:
                    break
                
        elif choice=='3':
            show(lists)
                
        elif choice=='4':
            while True:
                try:
                    ind=int(input("Enter place you want to delete: "))
                    delete(lists,ind)
                    break
                except ValueError:
                    break
            
        elif choice=='5':
            lists.clear()
            print("Successfully Deleted All")
            
        elif choice=='6':
            ask=input("Are you sure(yes,no): ")
            if ask=='yes':
                save(lists)
                print("Bye...")
                break
            else:
                continue
        
        else:
            print("Wrong Choice!")
    
if __name__==("__main__"):
    main()
