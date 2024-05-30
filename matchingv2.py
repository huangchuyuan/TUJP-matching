import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import time

# Function that prompts user to import a file
def prompt_file():
    """Create a Tk file dialog and cleanup when finished"""
    top = tk.Tk()
    top.wm_attributes("-topmost", 1)
    top.withdraw()  # hide window
    file_name = filedialog.askopenfilename(parent=top, title="Select a file", filetypes=[("Excel files", "*.xlsx"),("All files", "*.*")])
    top.destroy()
    return file_name

# Function that outputs imported files path for buddy and TUJP student files 
def import_files():
    has_imported_buddy = False
    while has_imported_buddy == False:
        print("Press enter to import buddy's response excel file")
        input1 = input()
        if input1 == "":
            buddy_file = prompt_file()
            if buddy_file == "":
                continue
            else:
                has_imported_buddy = True
        else:
            print("Invalid input, please press enter\n")
        
    has_imported_student = False
    while has_imported_student == False:
        print("Press enter to import TUJP student's response excel file")
        input2 = input()
        if input2 == "":
            student_file = prompt_file()
            if student_file == "":
                continue
            else:
                has_imported_student = True
        else:
            print("Invalid input, please press enter\n")

    return buddy_file, student_file


#_ = ts.preaccelerate_and_speedtest()

# Function that converts list to dict 
def list2dict(list, is_name=False):
    somedict = {}
    if is_name == True:
        for i in range(len(list)):
            somedict[i] = [str(list[i]).replace("\u3000","")]
    else:
        for i in range(len(list)):
            somedict[i] = str(list[i]).replace("\u3000","").lower().replace(" ","").split(",")
    
    return somedict



# Function that does the matching
def matching(buddy_file, student_file):

    print("Reading files...")
    # Read excel data 
    buddy_excel_data_df = pd.read_excel(buddy_file)
    student_excel_data_df = pd.read_excel(student_file)

    # Preprocessing of dataframe
    buddy_preprocessed_df = buddy_excel_data_df.rename(columns={"マッチングの参考のため、興味のある事柄についてお答えください。":"hobbies",
                                                    "マッチングの参考のため、第二外国語または学習中の言語についてお答えください。":"language",
                                                    "氏名":"name",
                                                    "性別":"gender",
                                                    "メールアドレス":"email",
                                                    "これまでのTUJPバディ経験 (回数を記入してください)":"tujptime"})
    student_preprocessed_df = student_excel_data_df.rename(columns={"メールアドレス":"email", "日本語レベル":"jplevel"})
    
    # Prepare dict for each category of matching form (for buddy and TUJP students)
    # buddy_languages = list2dict(buddy_preprocessed_df["language"])
    buddy_gender = list2dict(buddy_preprocessed_df["gender"])
    buddy_name = list2dict(buddy_preprocessed_df["name"], is_name=True)
    buddy_email = list2dict(buddy_preprocessed_df["email"])
    buddy_englevel = list2dict(buddy_preprocessed_df["englevel"])
    buddy_tujptime = list2dict(buddy_preprocessed_df["tujptime"])
    buddy_hobby = list2dict(buddy_preprocessed_df["hobbies"])

    student_hobbies = list2dict(student_preprocessed_df["Hobbies"])
    # student_languages = list2dict(student_preprocessed_df["language"])
    student_gender = list2dict(student_preprocessed_df["Gender"])
    student_name = list2dict(student_preprocessed_df["Name"],  is_name=True)
    student_email = list2dict(student_preprocessed_df["email"])
    student_jplevel = list2dict(student_preprocessed_df["jplevel"])

    print(int(buddy_englevel[0][0]))
    print(buddy_tujptime)
    print()

    print("Start matching...")
    # Prepare dict to store matched buddy and TUJP students
    matched = {}
    index_left = [x for x in range(len(student_name))]
    student_gender_left = list(student_preprocessed_df["Gender"])

    for i in range(len(buddy_name)):
        matched[i] = buddy_name[i]
        matched[i].append(buddy_email[i][0])

    # Start matching
    gender_weight = 5
    while len(index_left) > 0:
        for i in range(len(buddy_name)):
            max_count = -1

            for j in index_left:
                count = 0

                if buddy_gender[i] == student_gender[j]:
                    count += gender_weight
                # for item in buddy_hobbies[i]:
                #     if item in student_hobbies[j]:
                #         count += 1
                # for language in buddy_languages[i]:
                #     if language in student_languages[j]:
                #         count += 1
                
                if max_count < count:
                    max_count = count
                    best_match_index = j


            if len(index_left) == 0:
                break
            
            if buddy_gender[i][0] not in student_gender_left and count <= gender_weight:
                
                continue

            
            index_left.remove(best_match_index)
            student_gender_left.remove(student_gender[best_match_index][0])

            popped_student_name = student_name.pop(best_match_index)
            popped_student_email = student_email.pop(best_match_index)
            del student_gender[best_match_index]
            del student_hobbies[best_match_index]
            # del student_languages[best_match_index]
            

        
            matched[i].append(popped_student_name[0])
            matched[i].append(popped_student_email[0])

            

    with open('matching_results.csv', 'w', encoding="utf-8") as f:
        print("Buddy name, Buddy email, Student name and Student email", file=f)
        for i in range(len(matched)):    
            print(f"{",".join(matched[i])}", file=f)
            
    print("Outputted results in file: matching_results.csv")
    time.sleep(5)
    
# Main function            
def main():   
    # try:
        buddy_file, student_file = import_files()
        matching(buddy_file=buddy_file, student_file=student_file)
    # except Exception as e:
    #     print('Unexpected error:' + str(e))
    #     print("\nPress Ctrl + C to exit program")         
    #     time.sleep(60)    


if __name__ == "__main__":
    main()