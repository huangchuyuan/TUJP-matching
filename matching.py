import pandas as pd
import translators as ts

#_ = ts.preaccelerate_and_speedtest()

# Read excel data 
buddy_excel_data_df = pd.read_excel("2023  TUJP バディマッティング（回答） - Copy.xlsx", sheet_name="Form Responses 1")
student_excel_data_df = pd.read_excel("test_student_response.xlsx", sheet_name="Form Responses 1")

# Preprocessing of dataframe
buddy_preprocessed_df = buddy_excel_data_df.rename(columns={"マッチングの参考のため、興味のある事柄についてお答えください。":"hobbies",
                                                 "マッチングの参考のため、第二外国語または学習中の言語についてお答えください。":"language",
                                                 "氏名":"name",
                                                 "性別":"gender",
                                                 "メールアドレス":"email"})
student_preprocessed_df = student_excel_data_df.rename(columns={"マッチングの参考のため、興味のある事柄についてお答えください。":"hobbies",
                                                 "マッチングの参考のため、第二外国語または学習中の言語についてお答えください。":"language",
                                                 "氏名のアルファベット表記":"name",
                                                 "性別":"gender",
                                                 "メールアドレス":"email"})

# Function that converts list to dict 
def list2dict(list, translation, is_name):
    somedict = {}

    if translation == True:
        for i in range(len(list)):
            somedict[i] = ts.translate_text(str(list[i]).replace("\u3000",""),  translator="google", from_language="auto", to_language="en").lower().replace(" ","").split(",")
    elif translation == False and is_name == True:
        for i in range(len(list)):
            somedict[i] = [str(list[i]).replace("\u3000","").replace(","," ")]
    elif translation == False and is_name == False:
        for i in range(len(list)):
            somedict[i] = [str(list[i]).replace("\u3000","").lower()]
    
    return somedict


# Prepare dict for each category of matching form (for buddy and TUJP students)
buddy_hobbies = list2dict(buddy_preprocessed_df["hobbies"], translation=True, is_name=False)
buddy_languages = list2dict(buddy_preprocessed_df["language"], translation=True, is_name=False)
buddy_gender = list2dict(buddy_preprocessed_df["gender"], translation=True, is_name=False)
buddy_name = list2dict(buddy_preprocessed_df["name"], translation=False, is_name=True)
buddy_email = list2dict(buddy_preprocessed_df["email"], translation=False, is_name=False)

student_hobbies = list2dict(student_preprocessed_df["hobbies"], translation=True, is_name=False)
student_languages = list2dict(student_preprocessed_df["language"], translation=True, is_name=False)
student_gender = list2dict(student_preprocessed_df["gender"], translation=False, is_name=False)
student_name = list2dict(student_preprocessed_df["name"], translation=False, is_name=True)
student_email = list2dict(student_preprocessed_df["email"], translation=False, is_name=False)


# Function that checks if only one gender is remaining in the remaining unmatched TUJP student list
def check_gender(gender_list):
    if ("male" in gender_list or "Male" in gender_list) and ("female" in gender_list or "Female" in gender_list):
        return False
    else:
        return True


# Main function that does the matching
def matching():
    # Prepare dict to store matched buddy and TUJP students
    matched = {}
    index_left = [x for x in range(len(student_name))]
    student_gender_left = list(student_preprocessed_df["gender"])

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
                for item in buddy_hobbies[i]:
                    if item in student_hobbies[j]:
                        count += 1
                for language in buddy_languages[i]:
                    if language in student_languages[j]:
                        count += 1
                
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
            del student_languages[best_match_index]
            

        
            matched[i].append(popped_student_name[0])
            matched[i].append(popped_student_email[0])

            

    with open('testout.csv', 'w', encoding="utf-8") as f:
        print("Buddy name, Buddy email, Student name and Student email", file=f)
        for i in range(len(matched)):    
            print(f"{",".join(matched[i])}", file=f)
            
             



matching()
