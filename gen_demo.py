from docx import Document
import random
import os

os.makedirs("demo_files", exist_ok=True)

names = [
    "Tanishk Bansal", "Priya Sharma", "Rahul Verma", "Anjali Singh",
    "Rohit Mehta", "Sneha Gupta", "Arjun Yadav", "Pooja Agarwal",
    "Karan Joshi", "Divya Mishra", "Vikram Tiwari", "Neha Patel",
    "Amit Kumar", "Ritika Saxena", "Siddharth Nair", "Ananya Reddy",
    "Harsh Malhotra", "Simran Kaur", "Yash Bhatt", "Kavya Iyer"
]

filler_sentences = [
    "The candidate has successfully completed the online registration process.",
    "All documents have been verified by the admission committee.",
    "The institution reserves the right to cancel admission if documents are found invalid.",
    "Please report to the administrative office for further formalities.",
    "The academic session commences from the first week of August.",
    "Fee payment must be completed within seven days of admission.",
    "Students are required to maintain a minimum attendance of seventy five percent.",
    "The hostel allotment will be done on a first come first served basis.",
    "Identity verification is mandatory at the time of reporting.",
    "This document serves as proof of provisional admission to the institution.",
]

courses = ["B.Tech CSE", "B.Tech ECE", "B.Tech ME", "BCA", "MCA"]

for i in range(1, 31):
    doc = Document()
    name = random.choice(names)
    serial = f"STU-2024-{i:04d}"

    name_line = f"Name- {name}"
    serial_line = f"Serial No.- {serial}"

    fillers = random.sample(filler_sentences, random.randint(4, 8))

    paragraphs = fillers[:]
    pos1 = random.randint(0, len(paragraphs))
    paragraphs.insert(pos1, name_line)
    pos2 = random.randint(0, len(paragraphs))
    paragraphs.insert(pos2, serial_line)

    for para in paragraphs:
        doc.add_paragraph(para)

    doc.save(f"demo_files/student_{i:04d}.docx")

print("Done — 30 demo files created in demo_files/")
