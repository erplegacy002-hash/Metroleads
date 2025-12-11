// This file defines the mapping between the User (Assigned To) and the Project Site.
// Format: "User Name": "Site Name"

export const USER_PROJECT_MAPPING: Record<string, string> = {
  // Example mappings - Replace/Add as needed
  "Aishwarya Gulave": "Aqua life",
  "Raj Warde": "Aqua life",
  "Shubhantu Yadav": "milestone",
  "Smita Kad": "kairos",
  "Sanket Jejurkar": "kairos",
  "Pranav Satpute": "milestone",
  "Khushi Tamang": "milestone",
  "Rakshanda Gupta": "Aqua life",
  "Mohit Manani": "Aqua life",
  "Manisha Singh": "kairos",
  "Jai Mulik": "milestone",
  "Tanishq Singhai": "milestone",
  "Shubham Sangamnerkar": "milestone",
  "Sunil Mane": "milestone",
  "Omkar Khandge": "milestone",
  "Raunak Sharma": "milestone",
  "Sneha Patil": "milestone",
  "Neerja Sharma": "milestone",
  "Gauri Gokhale": "statement",
  "Shubham Pardesi": "Aqua life",
  "Vinita Bonde": "Aqua life",
  "Sonali Shinde": "kairos",
  "Sakshi Jamdar": "kairos",
  "Sejal Satav": "kairos",
  // Add more users here
};

// Default value if user is not found in the mapping
export const DEFAULT_SITE = "General Project";