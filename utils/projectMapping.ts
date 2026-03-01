// This file defines the mapping between the User (Assigned To) and the Project Site.
// Format: "User Name": "Site Name"

export const USER_PROJECT_MAPPING: Record<string, string> = {
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
  "Gauri Gokhale": "Kairos",
  "Shubham Pardesi": "Aqua Life",
  "Vinita Bonde": "Aqua Life",
  "Sonali Shinde": "Kairos",
  "Sakshi Jamdar": "Kairos",
  "Sejal Satav": "Statement", 
  "Rajshree Nimgire": "Kairos",
};

export const USER_TEAM_MAPPING: Record<string, string> = {
  "Aishwarya Gulave": "Sales",
  "Raj Warde": "Sales",
  "Shubhantu Yadav": "Sales",
  "Smita Kad": "Presales",
  "Sanket Jejurkar": "Sales",
  "Pranav Satpute": "Sales",
  "Khushi Tamang": "Sales",
  "Mohit Manani": "Sales",
  "Manisha Singh": "Presales",
  "Jai Mulik": "Sales",
  "Shubham Sangamnerkar": "Sales",
  "Sunil Mane": "Sales",
  "Omkar Khandge": "Sales",
  "Sneha Patil": "GRE",
  "Neerja Sharma": "GRE",
  "Gauri Gokhale": "Sales",
  "Shubham Pardesi": "Sales",
  "Sonali Shinde": "Sales",
  "Sakshi Jamdar": "GRE",
  "Sejal Satav": "Presales",
  "Rajshree Nimgire": "Sales Manager",
};

// Default value if user is not found in the mapping
export const DEFAULT_SITE = "General Project";
