import numpy as np
import pandas as pd
import os

################## Variables ####################################################################


################# DEFINE METHODS ################################################################
def convertMsg(testMsg, javadoc_score, format_score, grading_note, auto_score):
    manual_score = javadoc_score - format_score
        
    manual_part = f"""
**************************************************************\n
Manual Grading Score: {manual_score}/5\n
**************************************************************\n
Submission Format: {format_score}/5 (off points)\n
JavaDocs & Coding Conventions: {javadoc_score}/5(off points)\n
{grading_note}\n
**************************************************************\n
Early Submission: 0/5\n
WARNING: 0 token used.\n
**************************************************************\n
Automatic Grading Score: {auto_score}/95\n
**************************************************************\n"""
    finalMSG = ""
    array_of_sentences = testMsg.split('<br />')
    # Remove <b> from sentence 0 and 2
    array_of_sentences[0] = array_of_sentences[0].replace('<b>','').replace('</b>','')
    array_of_sentences[2] = array_of_sentences[2].replace('<b>','').replace('</b>','')
    
    # JOIN
    updatedMSG = '\n'.join(array_of_sentences).split('<br')[0]
    
    if (auto_score == 0):
        return manual_part + "\n" + updatedMSG
    
    # DO SOMETHING HERE
    updatedSentences = updatedMSG.replace('</ul>', '<ul>').split('<ul>')
    if len(updatedSentences) == 3:
        updatedSentences[1] = updatedSentences[1].replace('</li>', '<li>').split('<li>')
        updatedSentences[1] = ("\t".join(updatedSentences[1])).replace('\t\t', '\t')
        finalMSG = ("".join(updatedSentences))
    else:
        finalMSG = updatedSentences
    
    return manual_part + "\n" + finalMSG

#############################################################################################

# Load the grade file
source_df = pd.read_excel(os.listdir("manual")[0]) 

# Sort the file by last name and first name to match the order on Blackboard
source_df.sort_values(by=['Last Name', 'First Name'], ascending=True, inplace=True)
source_df = source_df.reset_index(drop=True)

# Overwrite the Feedback and total grade
for i in range(len(source_df.index)):
    try:
        source_df['Feedback to Learner'][i] = convertMsg(source_df['Feedback to Learner'][i], source_df['Javadoc'][i], source_df['Submission format'][i], source_df['Grading Notes'][i], source_df['Project 1 [Total Pts: 100 Score] |4045949'][i])
    except:
        print(i)
    
    source_df['Project 1 [Total Pts: 100 Score] |4045949'][i] = source_df['Project 1 [Total Pts: 100 Score] |4045949'][i] + source_df['Javadoc'][i] - source_df['Submission format'][i]
    if source_df['Project 1 [Total Pts: 100 Score] |4045949'][i] < 0:
        source_df['Project 1 [Total Pts: 100 Score] |4045949'][i] = 0

# Save the result into a new file
result_file = source_df.drop(['Javadoc', 'Submission format'], axis=1)
result_file.to_excel("output/feedbackFile.xls", index=False)  

# Load the blackboard file
bb_fileName = os.listdir("blackboard_downloaded")[0]
bb_df = pd.read_excel(bb_fileName)

total_point_column = result_file.columns[4]

# Merge two files by ID
df_cd = pd.merge(bb_df, result_file[['Student ID', total_point_column]], how='inner', on = 'Student ID')

# Override
df_cd[total_point_column + '_x'] = df_cd[total_point_column + '_y']
df_cd.drop([total_point_column + '_y'], axis=1, inplace=True)

# Rename columns
newName = []
for i in range(len(df_cd.columns)):
    if df_cd.columns[i][-2:] == "_x":
        newName.append(df_cd.columns[i][:-2])
    else:
        newName.append(df_cd.columns[i])

df_cd.set_axis(newName, axis=1, inplace=True)

# Save to a new file
df_cd.to_excel("output/upload_file.xls", index=False)