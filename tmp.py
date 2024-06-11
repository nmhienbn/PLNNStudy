import re
import csv
import random
import tkinter as tk
from tkinter import messagebox
import fitz  # PyMuPDF

# Step 1: Extract text from all pages of the PDF and parse into questions
def extract_text_from_all_pages(pdf_file):
    doc = fitz.open(pdf_file)
    text = ""
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        text += page.get_text("text")
    return text

def extract_questions(text):
    lines = text.split('\n')
    questions = []
    question = None
    choices = []
    correct_answers = []

    def add_question():
        if question and choices and correct_answers:
            questions.append({
                "question": question,
                "choices": choices,
                "correct_answers": correct_answers
            })

    for line in lines:
        line = line.strip()
        if line.startswith("Câu Hỏi"):
            add_question()
            question = line
            choices = []
            correct_answers = []
        elif re.match(r"^[a-hA-H]\.", line):
            choices.append(line[2:].strip())
        elif line.startswith("Câu trả lời đúng là:") or line.startswith("Đáp án chính xác là") or line.startswith("The correct answers are"):
            answers_text = line.split(":")[1].strip()
            correct_answers_text = [ans.strip() for ans in re.split(r'[,\s]', answers_text) if ans.strip()]
            correct_answers = []
            for ans in correct_answers_text:
                try:
                    correct_answers.append(choices.index(ans))
                except ValueError:
                    # Nếu câu trả lời không nằm trong danh sách, bỏ qua nó
                    pass
        elif line.startswith("Chọn câu:"):
            continue
        else:
            if choices and re.match(r"^[a-hA-H]\.", choices[-1]):
                choices[-1] += f" {line}"
            else:
                choices.append(line)

    add_question()
    return questions

def save_to_csv(questions, csv_file):
    with open(csv_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Question", "Choices", "Correct_Answers"])
        for q in questions:
            question = q['question']
            choices = q['choices']
            correct_answers = q['correct_answers']
            writer.writerow([question, ','.join(choices), ','.join(map(str, correct_answers))])

# Step 2: Create the Quiz App using Tkinter
class QuizApp:
    def __init__(self, root, questions):
        self.root = root
        self.root.title("Quiz Application")
        self.root.geometry("800x600")  # Fixed window size
        self.root.resizable(False, False)

        self.current_question = 0
        self.score = 0

        self.questions = questions

        self.question_label = tk.Label(
            self.root, text="", font=("Arial", 16), wraplength=700, justify="left"
        )
        self.question_label.pack(pady=20, anchor="w", padx=50)

        self.options_var = tk.StringVar()
        self.options_frame = tk.Frame(self.root)
        self.options_frame.pack(pady=10, anchor="w", padx=50)

        self.next_button = tk.Button(self.root, text="Next", command=self.next_question)
        self.next_button.pack(pady=20)

        self.start_quiz()

    def start_quiz(self):
        self.current_question = 0
        self.score = 0
        self.display_question()

    def display_question(self):
        question_data = self.questions[self.current_question]
        self.question_label.config(text=f"Q{self.current_question + 1}: {question_data['question']}")

        choices = question_data["choices"]
        correct_answers = question_data["correct_answers"]
        indexed_choices = list(enumerate(choices))
        random.shuffle(indexed_choices)
        shuffled_choices = [choice for _, choice in indexed_choices]
        self.correct_indices = [index for index, choice in indexed_choices if index in correct_answers]

        self.options_var.set(None)  # Ensure no choice is pre-selected
        for widget in self.options_frame.winfo_children():
            widget.destroy()

        self.option_buttons = []
        for i, choice in enumerate(shuffled_choices):
            rb = tk.Radiobutton(
                self.options_frame,
                text=f"{chr(65 + i)}. {choice}",
                variable=self.options_var,
                value=i,
                anchor="w",
                wraplength=600,
                justify="left"
            )
            rb.pack(anchor="w", pady=5)
            self.option_buttons.append(rb)

    def next_question(self):
        selected_option = self.options_var.get()
        if selected_option == "":
            messagebox.showwarning("No selection", "Please select an option before proceeding.")
            return

        if int(selected_option) in self.correct_indices:
            self.score += 1

        self.current_question += 1
        if self.current_question < len(self.questions):
            self.display_question()
        else:
            self.show_results()

    def show_results(self):
        messagebox.showinfo("Quiz Completed", f"Your score is {self.score}/{len(self.questions)}")
        self.root.destroy()

# Main function to run the application
if __name__ == "__main__":
    # Extract questions from PDF
    pdf_file_path = "sample.pdf"
    text = extract_text_from_all_pages(pdf_file_path)
    questions = extract_questions(text)

    # Save questions to CSV (optional)
    csv_file = "quiz.csv"
    save_to_csv(questions, csv_file)

    # Run the Quiz App
    root = tk.Tk()
    app = QuizApp(root, questions)
    root.mainloop()
