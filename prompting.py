import time
from openai import OpenAI
from datetime import datetime
import json

def get_custom_key_list(prompts):
    custom_keys = []
    keys_to_go_through = list(prompts.keys())
    for k in prompts.keys():
        custom_keys.append(k)  # Add the current key
    return custom_keys  

max_wait_time = 120  # seconds
prompts = {
    'prompt2_firstimpr': (
        "You are a senior trainee assessor at ORMIT Talent. Your task is to provide a first impression summary of a trainee:\n" 
        "1. Read the 'Context and Task Description' file and the Assessment Notes Candidate thoroughly to understand the task requirements.\n" 
        "2. Compile a brief, to-the-point first impression on how the trainee comes across from section '2. First Impression' in the context file.\n" 
        "3. Use only the first impressions found in the assessment notes for this compilation.\n" 
        "4. Provide the output as a concise summary, strictly limited to a single string containing the first impression, with no additional information or text."
    ),
    "prompt3_personality": (
        "You are a senior trainee assessor at ORMIT Talent. Your task is to create a concise, personal personality description for a trainee named 'Piet'.\n"
        "1. Review the 'Context and Task Description', 'Assessment Notes Candidate', and 'Personality Test Results' documents thoroughly to understand the requirements and desired communication style. Use the examples in the 'Personality Section Examples' document to understand the writing style we want. Do **not** use this document for the content of the section.\n"
        "2. Only if you find details about background information (e.g. the title of their previous studies) on the trainee, mention their previous activities very briefly and straightforwardly.\n"
        "3. Write a description that captures the trainee's main personality traits **using specific examples from the assessment notes of this trainee**. Focus on their interactions, contributions to team dynamics, and any standout moments that show both their strengths and areas for improvement, with specific examples.\n"
        "4. Use direct and simple language, avoiding complex words, unnecessary details, and overly enthusiastic tones. Steer clear of terms like 'demeanor', 'journey', 'promising path', 'shines through', or 'framework.' Keep it relatable and authentic, with a bit of wit.\n"
        "5. Provide balanced feedback that recognizes the trainee's strengths and outlines areas for development. Frame observations clearly, with a focus on actionable insights.\n"
        "6. Keep it engaging but to the point. Aim for a word count of 350 to 500 words. Your goal is to paint a focused picture of the trainee's personality based on the assessment feedback, avoiding formal language.\n"
        "7. Only return the Personality Description text, keeping it between 350-500 words."
    ),
    'prompt4_cogcap': (
        "1. Read the 'Context and Task Description' file and 'Capacity test results' file carefully to fully understand the task requirements.\n" 
        "2. Retrieve the cognitive capacity scores for the trainee, which include six specific scores.\n" 
        "3. Present the scores in the exact order and format as specified in the 'Context and Task Description' file.\n" 
        "4. Provide the output as a Python list in the following format: [score_general_ability, score_speed, score_accuracy, score_verbal, score_numerical, score_abstract].\n" 
        "5. Do not include any additional information or text; return only the specified list."
    ),
    'prompt5_language': (
        "1. Read the 'Context and Task Description' file and Assessment Notes Candidate document thoroughly to fully understand the task requirements.\n" 
        "2. Analyze the trainee’s language skills in Dutch, French, and English based on the Assessment Notes Candidate file.\n" 
        "3. If a label from the options 'A1', 'A2', 'B1', 'B2', 'C1' or 'C2' has been explicitly given in the assessment notes, use this label for that language. If no such label, apply the language label guide found in the 'Context and Task Description' file under heading '5. Language Skills' to assign appropriate labels to the analyzed language skills.\n" 
        "4. Ensure that the output is formatted exactly as specified in the context file.\n" 
        "5. Provide the output as a Python list in the following format: [label_Dutch, label_French, label_English].\n" 
        "6. Do not include any additional information, explanations, or text; return only the specified list."
    ),
    'prompt6a_conqual': (
        "1. Read the 'Context and Task Description' file, Assessment Notes Candidate, and Personality Test Results document thoroughly to fully understand the task requirements.\n"
        "2. Identify and collect 6 or 7 of the trainee's strongest qualities based on the assessment notes. Each quality should be clear, down-to-earth, and in simple language. Focus on short, practical descriptions of skills or behaviors, avoiding complex or formal words.\n"
        "3. Keep each statement under 10 words, focusing on clear, everyday language.\n"
        "4. Provide the output as a Python list in the following format: [first_quality, second_quality, third_quality, fourth_quality, fifth_quality, sixth_quality, seventh_quality].\n"
        "5. Do not include any additional information, explanations, or text; return only the specified list."
    ),
    'prompt6b_conimprov': (
        "1. Read the 'Context and Task Description' file, Assessment Notes Candidate, and Personality Test Results document thoroughly to fully understand the task requirements.\n"
        "2. Identify and collect 4 or 5 of the trainee's improvement/development points based on the assessment notes. Each development point should be clear, down-to-earth, and in simple language. Focus on short, practical descriptions of skills or behaviors, avoiding complex or formal words.\n"
        "3. Keep each statement under 10 words, focusing on clear, everyday language.\n"
        "4. Provide the output as a Python list in the following format: [first_improvement, second_improvement, third_improvement, fourth_improvement, fifth_improvement].\n"
        "5. Do not include any additional information, explanations, or text; return only the specified list."
    ),
    'prompt7_qualscore': (
        "1. Read through the Context file to understand the assessment.\n"
        "2. Identify and collect 6 or 7 of the strongest qualities of this trainee based on the assessment notes.\n"
        "3. Read through the MCP profile.\n"
        "4. For each of the previously established strengths, find the MCP profile subpoint that matches best with this strength.\n"
        "5. Return to me a list with zeros and ones, where only the subpoint matching each of the strengths best receives a 1 and the others receive a 0. For example, if Ownership was the only trait matching a strength best, the list looks like: [0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0].\n"
        "6. Make sure that the total list has 20 items, where each item corresponds to a specific subpoint in the MCP profile.\n"
        "7. Ensure that each strength is mapped to only **one** unique subpoint, and no subpoint should receive more than one `1`.\n"
        "8. Ensure that the list contains at least 5 and at most 7 `1`s, corresponding to the assessed strengths.\n"
        "9. **The output should contain only the specified list, with no additional text, explanation, or formatting.**\n"
        "10. **Do not include any text, comments, or explanation in the output—only the list of zeros and ones.**"
    ),
    'prompt7_improvscore': (
        "1. Read through the Context file to understand the assessment.\n"
        "2. Identify and collect 4 to 5 professional and/or personal improvement points for this trainee based on the assessment notes.\n"
        "3. Read through the MCP profile.\n"
        "4. For each of the previously established improvement points, find the MCP profile subpoint that matches best with this improvement point.\n"
        "5. Return to me a list with zeros and minus ones, where only the subpoint matching each of the improvement points best receives a -1 and the others receive a 0. For example, if Ownership was the only trait matching an improvement point best, the list looks like: [0,0,0,-1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0].\n"
        "6. Make sure that the total list has 20 items, where each item corresponds to a specific subpoint in the MCP profile.\n"
        "7. Ensure that each improvement point is mapped to only **one** unique subpoint, and no subpoint should receive more than one `-1`.\n"
        "8. Ensure that the list contains at least 3 and at most 5 `-1`s, corresponding to the assessed improvement points.\n"
        "9. **The output should contain only the specified list, with no additional text, explanation, or formatting.**\n"
        "10. **Do not include any text, comments, or explanation in the output—only the list of zeros and minus ones.**"
    ),
}

def send_prompts(data):
    print('Prompting started')
    results = {}
    #For filename:
    current_time = datetime.now()
    formatted_time = current_time.strftime("%m%d%H%M")  # Format as MMDDHHMinMin

    #Custom: Redacted anonymous files
    path_to_notes = 'temp/Assessment Notes.pdf'
    path_to_persontest = 'temp/PAPI Feedback.pdf'
    path_to_cogcap = 'temp/Cog. Test.pdf'
    #Default
    path_to_contextfile = 'resources/Context and Task Description.docx'
    path_to_toneofvoice = 'resources/Examples Personality Section.docx'
    path_to_mcpprofile = 'resources/The MCP Profile.docx'
    
    start_time = time.time()
    
    lst_files = [path_to_notes,
                 path_to_persontest,
                 path_to_cogcap,
                 path_to_contextfile,
                 path_to_toneofvoice,
                 path_to_mcpprofile
                 ]
    
    mykey = data["OpenAI Key"]
    
    # Initialize the OpenAI client with your API key
    client = OpenAI(api_key=mykey)
    
    # Create the assistant
    assistant = client.beta.assistants.create(
        name="ORMIT Report Assessor",
        instructions="You are a senior trainee assessor at a Belgian company ORMIT Talent. Your task is to extract and provide assessment data for a trainee based on notes from assessors who met the trainee and a personality and cognitive capacity test. You also have a context/task elaboration file and a file specifying the tone of voice for this.",
        model="gpt-4o-mini",
        tools=[{"type": "file_search"}],
    )
    
    # Ensure that the assistant has been created and you can access its ID
    if not hasattr(assistant, 'id'):
        raise Exception("Assistant creation failed or ID not found.")
    
    # Create the vector store
    vector_store = client.beta.vector_stores.create(name="Assessment_Data", expires_after={
        "anchor": "last_active_at",
        "days": 1
    })
    
    # Prepare files for upload
    file_streams = [open(path, "rb") for path in lst_files]
    
    # Upload files to the vector store
    file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
        vector_store_id=vector_store.id, files=file_streams
    )
    
    # Close file streams after upload
    for file in file_streams:
        file.close()
    
    print(file_batch.file_counts)
    
    # (Re)connect the assistant to the vector store(s)
    assistant = client.beta.assistants.update(
        assistant_id=assistant.id,
        tool_resources={"file_search": {"vector_store_ids": [vector_store.id]}},
    )
    
    # Now you should have access to the updated assistant's ID
    assistID = assistant.id
    print(f"Assistant ID: {assistID}")
    
    # Collect prompts
    lst_prompts = get_custom_key_list(prompts)
    print(lst_prompts)
    
    for prom in lst_prompts:
        print(prom)
        if prom in ['prompt4_cogcap', 'prompt6a_conqual']:
            time.sleep(61) #To not overload server and avoid minute rate limit
               
        # Create a new thread for each prompt
        empty_thread = client.beta.threads.create()
        
        # Create a message in the new thread
        client.beta.threads.messages.create(
            empty_thread.id,
            role="user",
            content=prompts[prom],
        )
        
        # Run the assistant
        run = client.beta.threads.runs.create(
            thread_id=empty_thread.id,
            assistant_id=assistID,  # Use assistID here
        )
        
        start_wait_time = time.time()

        while run.status != 'completed':
            time.sleep(2)  # Avoid overloading the server
            run = client.beta.threads.runs.retrieve(thread_id=empty_thread.id, run_id=run.id)
            
            # Check if wait time exceeds the max limit
            if time.time() - start_wait_time > max_wait_time:
                print(f"Timeout for {prom}")
                output = ''
                break
            # Only proceed if the run is completed
            if run.status == 'completed':
                output = client.beta.threads.messages.list(thread_id=empty_thread.id, run_id=run.id)
                messages = list(output)
                message_content = messages[0].content[0].text
                output = message_content.value
                            
        output_label = prom

        results[output_label] = output
        print(f"{output_label}: {output}")
        
        appl_name = data["Applicant Name"]
        filename_with_timestamp = f"data/{appl_name}_{formatted_time}.json"  # Filename with timestamp
        with open(filename_with_timestamp, 'w') as json_file:
            json.dump(results, json_file, indent=4)  # 'indent=4' for pretty printing
            
    # Clean vector store
    file_ids = {file.id for file in client.files.list()}
    print(file_ids)
    
    # Then attempt deletion only for files that are still in the list
    for file_id in file_ids:
        client.files.delete(file_id)
        print(f"Deleted file {file_id}")
    client.beta.assistants.delete(assistant.id)
    
    return filename_with_timestamp