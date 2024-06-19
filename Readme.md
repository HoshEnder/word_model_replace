# 📜 Custom Word Template Script 📜

Welcome to the Custom Word Template Script! This handy script allows you to easily replace Microsoft Word's default Normal template with your own custom template. Say goodbye to manually configuring your styles every time you start a new document! 🙌

## 🎨 Setting Up Your Custom Template

Before using the script, you need to create your custom Word template. Here's how:

1. Open a new Word document.
2. Customize the styles to your liking. This can include:
   - Modifying the default font, size, and color for the Normal style
   - Creating new styles for headings, subheadings, etc.
   - Setting up custom page margins, page size, and orientation
   - Defining your preferred paragraph spacing and indentation
   - ... and any other formatting you want to be applied by default! ✨
3. Once you're happy with your styles, save the document as a Word Template (`.dotm` file) with the name `normal.dotm`.
4. Place this `normal.dotm` file in the same folder as the `SetCustomTemplate.bat` script.

## 🚀 Using the Script

Now that you have your custom `normal.dotm` template ready, using the script is a breeze:

1. Double-click on `SetCustomTemplate.bat` to run it.
2. The script will:
   - Check if a `Normal.dotm` file already exists in Word's Templates folder
   - If it does, it will rename the existing file to `Normal.dotm_X`, where `X` is an incremental number (so your old templates are preserved) 📁
   - Copy your custom `normal.dotm` file into Word's Templates folder
3. That's it! Your custom template is now active. 🎉

## 💡 Tips

- If you want to update your custom template, just modify your `normal.dotm` file and run the script again. It will replace the old one while keeping a backup.
- If you need to revert to an older template, locate the `Normal.dotm_X` file you want in Word's Templates folder, delete the current `Normal.dotm`, and rename the `_X` file to `Normal.dotm`.

## 🤝 Contribution

Feel free to adapt and enhance this script to suit your needs. If you come up with a cool improvement, consider sharing it back with the community! 😊

---

Now go forth and enjoy your beautifully formatted Word documents, every time! 🌟

If you have any questions or suggestions, don't hesitate to reach out. Happy templating! 😄
