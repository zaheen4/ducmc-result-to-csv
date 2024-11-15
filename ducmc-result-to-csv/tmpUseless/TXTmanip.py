file_path = "Output/ResultFromWebsite.txt"
regi = 714

with open(file_path, "r") as file:
    content = file.read()

      # Replace all newline characters with custom delimeter for CSV
    delim = ";"
    content = content.replace("\n", delim)


    #######
    # Count the number of semicolons in the content
    semicolon_count = content.count(";")

    # Check if the semicolon count is exactly 3
    if semicolon_count == 3:
        # Split the content into parts
        parts = content.split(";")
        
        # Insert an extra semicolon after the 2nd semicolon
        modified_content = ";".join(parts[:2] + [""] + parts[2:])
    else:
        # Keep the content unchanged
        modified_content = content
    #######

    file_path_2 = "Output/ResultFromWebsite2.txt"
    with open(file_path_2, "w") as file:
        file.write(f"{modified_content}")

