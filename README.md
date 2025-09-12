\# Git Workflow Guide - Personal \& Office Development



> A comprehensive guide for managing Git repositories across personal laptop and office PC using Git Bash

> 

> \*\*Author:\*\* hasandafa  

> \*\*Last Updated:\*\* September 2025



\## Step 1: Create New Repository on GitHub



\### 1.1 Create Repository

1\. Go to \[GitHub.com](https://github.com) and sign in with your account

2\. Click the "+" icon in the top right corner and select "New repository"

3\. Set repository name: `your\_project\_name`

4\. Choose visibility (Public or Private)

5\. \*\*Do NOT\*\* initialize with README, .gitignore, or license (if you have existing local files)

6\. Click "Create repository"



\### 1.2 Initialize Local Repository (First Time Setup)

Open Git Bash in your project directory and run:

```

\# Initialize git repository

git init



\# Add all files

git add .



\# Create initial commit

git commit -m "Initial commit: Project setup"



\# Rename default branch from master to main

git branch -m master main



\# Add GitHub repository as remote origin

git remote add origin https://github.com/user.name/your\_project\_name.git



\# Push to GitHub and set main as upstream branch

git push -u origin main

```


\### 1.3 Configure Git Default Branch (One-time setup)

'''
# Set main as default branch for all future repositories

git config --global init.defaultBranch main



\# Set your Git credentials (if not done before)

git config --global user.name "user.name"

git config --global user.email "emai@example.com"
'''



\## Step 2: Access Repository from Another Device



\### 2.1 Clone Repository (Office PC or New Device)

```

\# Navigate to your workspace directory

cd /c/path/to/your/workspace



\# Clone the repository

git clone https://github.com/hasandafa/your\_project\_name.git



\# Navigate into the cloned directory

cd your\_project\_name

'''



\### 2.2 Verify Setup

'''

\# Verify you're on main branch

git branch

# Check current branch

git branch



\# Check remote configuration

git remote -v



\# Check branch tracking

git branch -vv

'''



\## Step 3: Daily Workflow - Add/Update Programs or Files

\### 3.1 Before Starting Work (Always Do This First)

'''

\# Pull latest changes from GitHub

git pull origin main

'''



\### 3.2 Working with FilesS
'''

\# Check current status

git status



\# Add specific file(s)

git add filename.py

git add folder/



\# Add all changed files

git add .



\# Check what will be committed

git status

'''



\### 3.3 Commit changes

'''

\# Commit with descriptive message

git commit -m "Add feature: user authentication"

git commit -m "Fix bug: form validation error"

git commit -m "Update: improve database connection"



\# Or commit with detailed description

git commit -m "Add user registration form



\- Create registration form with validation

\- Add password confirmation field

\- Implement email verification"

'''



\### 3.4 Push Changes to GitHub

'''

\# Push to main branch

git push origin main



\# Or simply (after first push with -u)

git push

'''



\## Step 4: Advanced Git Operations

\### 4.1 Check Project History

'''

\# View commit history

git log



\# View compact history

git log --oneline



\# View specific file history

git log filename.py

'''



\### 4.2 Undo Changes

'''

\# Undo changes in working directory (before add)

git checkout -- filename.py



\# Unstage file (after add, before commit)

git reset HEAD filename.py



\# Undo last commit (keep changes in working directory)

git reset HEAD~1



\# Undo last commit (discard all changes)

git reset --hard HEAD~1

'''



\### 4.3 Branch Management

'''

\# Create and switch to new branch

git checkout -b feature-branch



\# Switch between branches

git checkout main

git checkout feature-branch



\# List all branches

git branch



\# Merge feature branch to main

git checkout main

git merge feature-branch



\# Delete branch after merging

git branch -d feature-branch

'''



\### 4.4 Handle Conflicts

'''

\# When pull fails due to conflicts

git pull origin main



\# Edit conflicted files manually, then:

git add .

git commit -m "Resolve merge conflicts"

git push origin main

'''



\## Step 5: Troubleshooting Common Issues

\### 5.1 Authentication Issues (Enterprise Network)

'''

\# If using SSH keys

git remote set-url origin git@github.com:hasandafa/your\_project\_name.git



\# If using HTTPS with token

git remote set-url origin https://username:token@github.com/hasandafa/your\_project\_name.git

'''



\### 5.2 Sync Issues Between Devices

'''

\# Force pull (be careful, this overwrites local changes)

git fetch origin

git reset --hard origin/main



\# Create backup before force pull

git stash

git pull origin main

git stash pop

'''



\### 5.3 Check Repository Status

'''

\# Detailed status information

git status -v



\# Check differences

git diff



\# Check differences of staged files

git diff --staged



\# Check remote repository info

git remote show origin

'''



\## Step 6: Best Practices

\### 6.1 Commit Messages Guidelines

\- Use present tense: "Add feature" not "Added feature"

\- Keep first line under 50 characters

\- Use descriptive messages

\- Examples:

&nbsp; - `Add user authentication system`

&nbsp; - `Fix database connection timeout`

&nbsp; - `Update API documentation`

&nbsp; - `Refactor code structure`



\### 6.2 File Management

'''

\# Create .gitignore file to exclude unnecessary files

echo "\_\_pycache\_\_/" >> .gitignore

echo "\*.pyc" >> .gitignore

echo ".env" >> .gitignore

echo "node\_modules/" >> .gitignore



\# Add .gitignore to repository

git add .gitignore

git commit -m "Add .gitignore file"

'''



\### 6.3 Regular Workflow Checklist

1. ✅ '''git pull origin main''' (before starting work)

2\. ✅ Make your changes

3\. ✅ '''git add .''' (stage changes)

4\. ✅ '''git commit -m "descriptive message"''' (commit changes)

5\. ✅ '''git push origin main''' (push to GitHub)



\## Quick Reference Commands

'''

\# Essential daily commands

git pull origin main           # Get latest changes

git add .                      # Stage all changes

git commit -m "message"        # Commit changes

git push origin main           # Push to GitHub

git status                     # Check current status

git log --oneline             # View commit history



\# Setup commands (one-time)

git init                      # Initialize repository

git remote add origin URL     # Add remote repository

git branch -m master main     # Rename branch to main

git push -u origin main       # First push with upstream



\# Troubleshooting

git stash                     # Temporarily save changes

git stash pop                 # Restore stashed changes

git reset --hard origin/main  # Reset to remote version

'''

