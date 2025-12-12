// This file defines the mapping between the User (Assigned To) and the Project Site.
// Format: "User Name": "Site Name"

export const USER_PROJECT_MAPPING: Record<string, string> = {
  // Example mappings - Replace/Add as needed
  "Aishwarya Gulave": "Aqua Life",
  "Raj Warde": "Aqua Life",
  "Shubhantu Yadav": "Milestone",
  "Smita Kad": "Kairos",
  "Sanket Jejurkar": "Kairos",
  "Pranav Satpute": "Milestone",
  "Khushi Tamang": "Statement",
  "Rakshanda Gupta": "Aqua Life",
  "Mohit Manani": "Aqua Life",
  "Manisha Singh": "Kairos",
  "Jai Mulik": "Milestone",
  "Tanishq Singhai": "Milestone",
  "Shubham Sangamnerkar": "Milestone",
  "Sunil Mane": "Milestone",
  "Omkar Khandge": "Milestone",
  "Raunak Sharma": "Milestone",
  "Sneha Patil": "Milestone",
  "Neerja Sharma": "Milestone",
  "Gauri Gokhale": "Statement",
  "Shubham Pardesi": "Aqua Life",
  "Vinita Bonde": "Aqua Life",
  "Sonali Shinde": "Kairos",
  "Sakshi Jamdar": "Kairos",
  "Sejal Satav": "Statement", 
  // Add more users here
};

// Default value if user is not found in the mapping
export const DEFAULT_SITE = "General Project";
