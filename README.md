# How to Use AutoRepairForm.exe

## Option 1: Use the Existing Program File

You can open the program file from the shared folder, or download it from the attached email.

Press:

```text
Windows + R
````

Then enter the shared folder path:

```text
\\10.150.208.54\SF Data Files 2010\AE\thitiporn\Scrap
```

![Open shared folder](https://github.com/user-attachments/assets/47260396-337d-483c-bb6c-6f9d43814ab3)

---

## Option 2: Clone the Git Branch and Build the Program

This option is for users who need to get the source code and build the `.exe` file manually.

### Step 1: Clone the Git Repository

Open **Command Prompt**, **Git Bash**, or **Terminal**.

Run the command below:

```bash
git clone -b scrapfilledform <repository-url>
```

Example:

```bash
git clone -b scrapfilledform https://github.com/your-username/your-repository-name.git
```

> **Note:**
> Replace `<repository-url>` with the real Git repository URL.

---

### Step 2: Go to the Project Folder

After cloning is complete, go into the project folder:

```bash
cd your-repository-name
```

---

### Step 3: Install Required Packages

Run:

```bash
pip install -r requirements.txt
```

> **Note:**
> Make sure Python is already installed on your computer.

---

### Step 4: Build the Program to `.exe`

Run the build command:

```bash
pyinstaller --onefile --windowed AutoRepairForm.py
```

Or, if the project already has a `.spec` file, run:

```bash
pyinstaller AutoRepairForm.spec
```

---

### Step 5: Find the `.exe` File

After the build is complete, the `.exe` file will be created in the `dist` folder.

Example:

```text
dist/AutoRepairForm.exe
```

You can copy this `.exe` file to the shared folder or send it by email.

---

## Step 1: Run the Program

Double-click **AutoRepairForm.exe**.

![AutoRepairForm.exe](https://github.com/user-attachments/assets/cac1901b-10a1-4b14-ada7-4a7d8c62c7aa)

---

## Step 2: Allow the Program to Run

If Windows shows a security warning:

1. Click **More info**
2. Click **Run anyway**

![Windows security warning](https://github.com/user-attachments/assets/35d51d8b-667f-41d1-9eb2-8e17e5a7fd1f)

---

## Step 3: Enter or Import Serial Numbers

You can enter the serial number in one of two ways:

* Enter one serial number manually in the input box
* Click **Import Excel** to import multiple serial numbers from an Excel file

![Enter or import serial number](https://github.com/user-attachments/assets/a90e53c0-dba3-4619-a9aa-30de129a08f2)

---

## Step 4: Check the Configuration

After entering or importing the serial numbers, the configuration form will appear.

Please check the default configuration.
You can adjust the configuration if needed.

> **Note:**
> Before clicking **Run**, make sure the **Repair-Rev** page is already opened at the **Input Defect SN** page.

![Configuration form](https://github.com/user-attachments/assets/4f43ec9d-05a2-421c-9589-b4621c0b6d84)

![Repair-Rev Input Defect SN page](https://github.com/user-attachments/assets/5cc894d1-1377-47f2-a994-68247dc9cbf2)

---

## Step 5: Run the Program

After confirming the configuration and making sure the Repair-Rev page is ready, click **Run**.

The program will start filling in the serial number information automatically.

````

For your sentence, you can write it like this:

```text
You can also clone the Git branch named `scrapfilledform` and build the program into an `.exe` file manually.
````
