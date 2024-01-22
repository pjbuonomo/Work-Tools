library(reticulate)

# Install the 'transformers' and 'torch' library if not available
if (!py_module_available("transformers")) {
  py_install("transformers")
}
if (!py_module_available("torch")) {
  py_install("torch")
}
if (!py_module_available("tensorflow")) {
  py_install("tensorflow")
}

# Load the Python library for transformers
transformers <- import("transformers")
tensorflow <- import("tensorflow")

# Define the model name
model_name_or_path <- "bert-base-uncased"

# Load the pre-trained model and tokenizer
tokenizer <- transformers$BertTokenizer$from_pretrained(model_name_or_path)
model <- transformers$TFAutoModel$from_pretrained(model_name_or_path)

# Function to classify and format lines
classify_and_format <- function(text, your_classification_function, your_bid_value) {
  # Tokenize and encode the text
  inputs <- tokenizer$encode_plus(text, return_tensors="tf")
  
  # Pass the input through the model to get embeddings
  outputs <- model(inputs)
  embeddings <- outputs$last_hidden_state
  
  # Example: Classification logic
  # Replace this with your actual classification logic
  classification_result <- your_classification_function(embeddings)
  
  if (classification_result == "Size") {
    # Format for "Size" lines
    formatted_output <- paste(text, "bid @ ", your_bid_value, sep="")
  } else {
    # Format for "Name" lines
    formatted_output <- paste(text, "bid @ ", your_bid_value, sep="")
  }
  
  return(formatted_output)
}

# Example usage (assuming 'your_classification_function' and 'your_bid_value' are defined)
email_content <- c(
  "4.25mm Cosaint 2021-1 A (22112CAA0) 99.70 bid / 100.10 offer",
  "Kilimanjaro III Re 2021-2 B-2 (49407PAJ9) bid at 95.60",
  "Matterhorn Re 2022-1 A (577092AP4) bid at 97.75",
  "Mystic Re IV 2023-1 A (62865LAD9) bid at 102.90")

#Example: Define a placeholder function for classification and bid value
your_classification_function <- function(embeddings) {
  
#Placeholder logic for classification
  return("Size") # or "Name", depending on your logic
}
your_bid_value <- "100" # Placeholder bid value

for (line in email_content) {
  formatted_line <- classify_and_format(line, your_classification_function, your_bid_value)
  cat(formatted_line, "\n")
}

#To print embeddings, select a specific text and run the model
selected_text <- email_content[1]
inputs <- tokenizer$encode_plus(selected_text, return_tensors="tf")
outputs <- model(inputs)
embeddings <- outputs$last_hidden_state
print(embeddings)

