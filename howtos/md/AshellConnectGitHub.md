# Title: a-shell connect GitHub

#### Tags: A-shell, GitHub, HowTo, Instructions

Intro: I have had a hard time in the past trying to make .txt files and being able to work on simple files on GitHub
repos from my phone and iPad. There are times when I wanted to just write a quick journal entry or work on a file on my
repo while I was waiting for a friend to show up for a lunch hangout, or when I was just bored waiting. I am excited
that
this exists.

A-Shell is an app in the App Store that I found recently due to someone saying that they only had their iPad for a
workshop on GitHub. In the instructions of the workshop, we ask them to use a computer, but this happens more often than
not. So I went on the search for a GitHub iPad GUI. At first, I found one that looked amazing, it had all the terms and
did almost all the same items as GitHub Desktop, but they charged you to push... I was so disappointed, then I found
another one that let you do 3 pushes for free, then you had to pay. I was slightly less disappointed. Yet this got me
thinking and searching to see if I can do it with the command line for free. Which is when I found A-shell, this is
where the headache started for a different reason. I needed to set up the connection to pull and push to GitHub. I was
unaware of
what my next 8 hours would look like.

Side note scene set up time (This is for nothing about a-shell, I just wanted to write this): It was a Friday night, my
cousin was having a small birthday party for her now 7-year-old child. She decided to hire a magician to do some magic,
which actually entertained not just the small children but all of the adults there. There I was sitting on my phone half
looking up at each magic trick, trying to connect a-shell to GitHub, and as I am silently cussing to myself, "I did make
it, why is it not working?" I keep getting the same error: the .ssh does not exist, when I did the cat copy code to put
it into GitHub's settings.

So once I was at the house, I was able to really research what was going on, which is when I found the forum post (I did
try ChatGPT by the way, ChatGPT tells you just go to the .ssh folder or here is the code, it will work. Whereas a forum
post will say, dude, I know I went through the same struggles, here is everything I have tried, and this is how I fixed
it. This is why I will always appreciate forums more than a chatbot.) that told me what I was missing, which is exactly
where the .ssh folder is on the iPad, when the pwd is semi lying to you... (it was not, but it was not picking it up
with the CAT code, and when you went looking for it, it was like nope, I am not here, until after 8 hours, it was like "
AHA YOU FOUND ME!") because you had just decided to write the same cd, then press tab not in the root folder, but in the
folder you were in the entire time. So, below are the steps for future me, so I do not forget.

## Steps:

1. Open A-shell
2. install git

```bash
    >pkg install git
```

3. Check with ls that there is a bin folder

```bash
    >ls
```

4. Also, check what commands you can do with lg2

```bash
    > lg2
```

5. Clone the git repo from GitHub. You will need the URL of the repo you are cloning. There is no big green button on
   the phone app, so go to repo, click the circle with the three dots, and click share, then just click the option copy.
   You will have to add ".git" at the end.

```bash 
    > git clone "GitHub Repo URL with .git at the end"
```

6. Check with ls that your repo is there

```bash
    > ls
```

7. Now generate the public/private key

```bash
    > ssh-keygen -t rsa -b 4096 -C "your email you use on GitHub"
```

8. Click enter for default file name and file folder, unless one already exists, click Y first, then for the passwords,
   click enter, then enter (You can make your first protected with passcodes, I did...) It will give you the path it's
   saved at for both private and public, and display the fingerprint, text, and graphic.

9. In the folder where you have the bin and the git clone, write cd, then hit tab, then type a dot, then "s" the .ssh/
   should come up if it doesn't, you will probably have to have an extra set of fun searching the answer for next 8
   hours, I recommend going to a magic show first. But it should have come up; I have tested it many times. Go into that
   folder

```bash
    > cd .ssh/
```

10. Check that the keys are in there with ls

```bash
    > ls
```

11. display the public key and manually copy it

```bash
    > head id_ed25519.pub
```

12. Go to GitHub on a browser (I used Firefox on my iPad and phone), navigate to the settings, then to SSH and GPG keys,
    click the green button that says "New SSH key", call it the device, and then paste the manually copied public key.
    This is a good point to make a GPG key, you will need it when you git push to the repo.

13. Make the GPG key, IMPORTANT: IT WILL ONLY DISPLAY THE TEXT IT GIVES YOU FOR THE GPG KEY ONCE. MAKE SURE TO COPY AND
    SAVE IT SOMEWHERE WHERE IT IS SAFE AND ACCESSIBLE. I have a friend who told me it was a pain replacing the old one.
    You do not want to go through that, right?

14. Okay, finally! Let's test it! Navigate to the git repo folder and make sure you are inside it.

```bash
    > cd "repo folder"
```

15. Make a test file, we will push it to the repo

```bash
    > vim testfile.txt
```

16. Pull the git repo

```bash
    > git pull
```

17. Add the file to be uploaded to git

```bash
    > git add .
```

18. Look at the status; it should say changes to commit or something similar

```bash
    > git status
```

19. make a commit message (whisper: you will not be able to do this, it will say you will have an extra step, the first
    time you do this, thank goodness for testing)

```bash
    > git commit -m "Your message here"
```

20. Oh, look, an error, I would never have guessed. Add your user.name and user.email from GitHub.

```bash
    > lg2 config user.name 'Your Name'
```

21. Add your user.email from GitHub now

```bash
    > lg2 config user.email youremail@example.com
```

22. Redo the commit message

```bash
    > git commit -m "Your message here"
```

23. Check the status, I only get back what branch it's on.

```bash
    > git status
```

24. Now push to the repo

```bash
    > git push
```

25. It will ask for the password; this is going to be the GPG key, not your GitHub password. I just copied and pasted it
    in there. (I tried my GitHub password for a good 10 attempts before I remembered to use the GPG key (if it gives you
    a error that says "Error Pushing [-1] - unexpected EOF", they go buy a ticket to the local magic show... it will be
    another 8 hours)

26. Success *chit's voice* (TikTok reference), if it says pushed. If, for some odd reason, you get the error "Error
    Pushing [-1] - unexpected EOF" or something else as an error... do you like card tricks? I know a magician you might
    like. Overplayed joke... yeah, I know, but you now have a inside joke with me (or yourself). But go check that it is
    on your GitHub Repo Online.

Summary: Well, that is how you connect a-shell to your GitHub. To future me, I know you are going to want to thank past
you (me) for taking the time to write this out, buy me a coffee or lunch if you want.

Time it takes to set up if successful: 5-10 minutes.

Links to the forums that I found useful:
"https://forum.obsidian.md/t/mobile-automatic-sync-with-github-on-ios-for-free-via-a-shell/46150?page=2"
it has links to other websites including the main GitHub Documentation on SSH, it is for Obsindian app, but it gave me
the information to fix my issue.

