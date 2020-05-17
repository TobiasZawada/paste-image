;;; paste-image.el --- Paste images from WSL-clipboard via powershell.  -*- lexical-binding: t; -*-

;; Copyright (C) 2018  Tobias Zawada

;; Author: Tobias Zawada <naehring@smtp.1und1.de>
;; Keywords: unix, multimedia

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; 

;;; Code:

(require 'subr-x)
(require 'image)

(defun paste-image-xselect-convert-to-png (_selection type value)
  "Image/png is stored as TYPE image/png with the image data string as VALUE.
This function just returns VALUE."
  (when (string-equal (substring value nil 8) ".PNG\x0D\x0A\x1A\x0A")
      (propertize value 'image-type type)))

(defun paste-image-xselect-convert-to-svg (_selection type value)
  "Image/svg is stored as TYPE image/svg with the image data string as VALUE.
This function just returns VALUE."
  (propertize value 'image-type type))

(defun paste-image-add-to-alist (list car cdr)
  "In LIST add the cons (CAR . CDR).
Do not add the cons if there is already an entry with key CAR
but issue a warning."
  (let ((old-cons (assoc car (symbol-value list))))
    (if (and old-cons
	     (null (eq (cdr old-cons) cdr)))
	(warn "Key %s is already in list %s" car list)
      (add-to-list list (cons car cdr)))))

(paste-image-add-to-alist 'selection-converter-alist 'image/png #'paste-image-xselect-convert-to-png)

(defun paste-image-function-default (data &optional type)
  "Insert a space with display property 'image and image DATA with TYPE at point."
  (insert (propertize " " 'display (create-image data type t))))

(defvar-local paste-image-function #'paste-image-function-default
  "Function that pastes the image from the clipboard at point.
The function should take the image data as its only argument.")

(defun paste-image ()
  "Paste image from clipboard at point."
  (interactive)
  (when-let
      ((data (cond
	      ((memq system-type '(windows-nt cygwin))
	       (let ((selection-coding-system 'no-conversion))
		 (gui-get-selection nil 'image/png)
		 ))
	      ((wsl-p)
	       (when-let* ((default-directory temporary-file-directory)
			   (tmp-file (make-temp-name (expand-file-name "clipboardImage" wsl-exchange-dir)))
			   (wsl-path (wsl-path "-w" "-a" tmp-file)))
		 (with-current-buffer (get-buffer-create "*powershell*")
		   (erase-buffer)
		   (call-process "powershell.exe" nil t nil
				 "-Command" (format "(Get-Clipboard -Format Image).Save(\"%s\", [System.Drawing.Imaging.ImageFormat]::Png)" wsl-path)))
		 (with-temp-buffer
		   (set-buffer-file-coding-system 'no-conversion)
		   (set-buffer-multibyte nil)
		   (insert-file-contents-literally tmp-file nil nil nil t)
		   (buffer-string)))))))
    (funcall paste-image-function data)))

(defun paste-image-to-clipboard (&optional image)
  "Copy IMAGE data to Windows clipboard."
  (interactive)
  (unless image
    (setq image (or (cl-loop for ol in (overlays-at (point))
			     when (overlay-get ol 'display)
			     return (overlay-get ol 'display))
		    (get-text-property (point) 'displqay)))
    (unless (consp image)
      (user-error "No image at point")))
  (when (and (consp image)
	     (eq (car image) 'image))
    (setq image (cdr image))
    (cond
     ((plist-member image :file)
      (setq image (plist-get image :file))
      (unless (file-name-absolute-p image)
	(setq image (expand-file-name image data-directory)))
      (setq image (with-temp-buffer
		    (set-buffer-multibyte nil)
		    (insert-file-contents-literally image)
		    (buffer-string))))
     ((plist-member image :data)
      (setq image (plist-get image :data)))))
  (unless (stringp image)
    (user-error "Wrong image format"))
  (with-current-buffer (get-buffer-create "*powershell*")
    (erase-buffer)
    (let* ((tmp-file (make-temp-name (expand-file-name "clipboardImage" wsl-exchange-dir)))
	   (wsl-path (wsl-path "-w" "-a" tmp-file))
	   (script (format "$FILE=Get-Item -Path '%s';[Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');[Reflection.Assembly]::LoadWithPartialName('System.Drawing');[System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile($FILE));" wsl-path)))
      (with-temp-file tmp-file
	(set-buffer-multibyte nil)
	(insert image))
      (unless (eq 0
		  (call-process
		   "powershell.exe" nil t nil
		   "-Command" script))
	(user-error "Powershell failed to copy image to clipboard")))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; Pasting image files

(defvar paste-image-file-hook nil "Run after pasting images from clipboard.")
(make-variable-buffer-local 'paste-image-file-hook)

(defvar paste-image-prefix "" "Prefix to be inserted before the pasted image.")
(make-variable-buffer-local 'paste-image-prefix)

(defvar paste-image-dir-postfix "" "Postfix to be inserted at end of image directory.")
(make-variable-buffer-local 'paste-image-dir-postfix)

(defvar paste-image-filename-prefix "img" "Prefix to be inserted before the pasted image.")
(make-variable-buffer-local 'paste-image-filename-prefix)

(defvar paste-image-postfix "" "Postfix to be inserted after the pasted image.")
(make-variable-buffer-local 'paste-image-postfix)

(defvar paste-image-type "png" "Default image type and extension of pasted images.")
(make-variable-buffer-local 'paste-image-type)

(unless (assoc "\\.eps" image-type-file-name-regexps)
  (add-to-list 'image-type-file-name-regexps '("\\.eps" . postscript)))

(defvar paste-image-types
  '(
    ("png" nil nil '(image-mode))
    ("eps" "png" "convert %< -resample 100 -compress lzw ps3:%>" '(doc-view-mode))
    )
  "Each type specifier is a list.  nth 0 is the target image format.
nth 1 is the source image format,
nth 2 is the converter from target to source
      %< stands for the source and
      %> stands for the output
nth 3 are Lisp commands with which the image can be displayed.")

(defun paste-image-type-from-file-name (name)
  "Return image type of file with NAME for paste-image-file."
  (car (assoc (downcase (file-name-extension name)) paste-image-types)))

(defun paste-image-paste (dir-name target-name img-type)
  ""
  (let*
      ((img-type-list (assoc img-type paste-image-types))
       (grab-type (or (nth 1 img-type-list) (car img-type-list)))
       (converter (nth 2 img-type-list))
       (src-name target-name)
       )
    (unless grab-type
      (error "Unsupported image grab type %s.  (See paste-image-types.)" img-type))
    (when converter
      (setq src-name (concat target-name "." grab-type))
      )
    (with-temp-buffer
      (setq default-directory dir-name)
      (insert "
from PIL import Image, ImageGrab
print \"--- Grab Image From Clipboard ---\"
im=ImageGrab.grabclipboard()
if isinstance(im,Image.Image):
 im.save(\"" src-name "\",\"" (upcase grab-type) "\")
else:
 print \"paste-image-paste:No image in clipboard.\"
")
      (call-process-region (point-min) (point-max) pythonw-interpreter t t nil)
      (goto-char (point-min))
      (if (search-forward "paste-image-paste:No image in clipboard." nil 'noError)
	  (error "No image in clipboard"))
      (message "%s" (buffer-substring-no-properties (point-min) (point-max)))
      (when converter
	(setq converter (replace-regexp-in-string "%<" src-name converter t t))
	(setq converter (replace-regexp-in-string "%>" target-name converter t t))
	(shell-command converter)
	(unless (equal src-name target-name)
	  (delete-file src-name))
	))
    ))

(defun paste-image-absolute-dir-name ()
  "Create an absolute dir name from return value of function `buffer-file-name'."
  (let ((bufi (buffer-file-name)))
    (concat (if bufi (file-name-sans-extension bufi) default-directory) paste-image-dir-postfix)))

(defun paste-image-relative-dir-name ()
  "Create a relative dir name from `paste-image-absolute-dir-name'."
  (file-relative-name (paste-image-absolute-dir-name) default-directory))

(defun paste-image-cleanup ()
  "Remove pasted image files not contained in the document anymore."
  (interactive)
  (let* ((re-images (concat  paste-image-filename-prefix "[0-9]*\\." paste-image-type))
	 (dir-name (paste-image-relative-dir-name))
	 (del-list (directory-files dir-name nil (concat "^" re-images "$") 'nosort)))
    (save-excursion
      (goto-char (point-min))
      (let ((re (concat dir-name "/\\(" re-images "\\)")))
	(while (search-forward-regexp re nil 'noErr)
	  (setq del-list (cl-delete (match-string-no-properties 1) del-list :test 'string-equal))
	  ))
      (cl-loop for f in del-list do
	    (let ((fullpath (concat dir-name "/" f)))
	      (with-temp-buffer
		(insert-file-contents fullpath)
		(set-buffer-modified-p nil)
		(eval (nth 3 (assoc paste-image-type paste-image-types)))
		(display-buffer (current-buffer))
		(when (yes-or-no-p (concat "Delete file " fullpath))
		  (delete-file fullpath)))))
      )))

(defun paste-image-new-file-name (prefix suffix)
  "Get name of non-existing file by inserting numbers between PREFIX and SUFFIX if necessary.
SUFFIX may not include directory components."
  (let ((first-try (concat prefix suffix))
	(prefix-dir (or (file-name-directory prefix) "./"))
	(prefix-file (file-name-nondirectory prefix)))
    (if (file-exists-p first-try)
	(concat
	 prefix
	 (number-to-string
	  (1+
	   (apply
	    'max
	    (append '(-1)
		    (mapcar #'(lambda (name)
				(string-to-number (substring name (length prefix-file) (- (length suffix)))))
			    (directory-files prefix-dir nil (concat prefix-file "[0-9]+" suffix) 'NOSORT))))))
	 suffix)
      first-try)))

(defun paste-image-file (&optional file-name img-type)
  "Paste image file from clipboard into file FILE-NAME with IMG-TYPE.
That means put the image file into the directory
with the basename of the buffer file.
The image gets a name 'imgXXX.png' where XXX stands for some unique number."
  (interactive)
  (unless img-type
    (setq img-type paste-image-type))
  (unless file-name
    (let* ((dir-name (paste-image-absolute-dir-name))
	   (file-path (paste-image-new-file-name (expand-file-name paste-image-filename-prefix dir-name) (concat "." img-type))))
      (setq dir-name (file-name-directory file-path))
      (setq file-name (file-name-nondirectory file-path))
      (while
	  (progn
	    (setq file-path (read-file-name "Image file name:" dir-name nil nil file-name))
	    (let ((dir (file-name-directory file-path)))
	      (or (and
		   dir
		   (file-exists-p dir)
		   (null (file-directory-p dir)))
		  (null (and (file-name-extension file-path)
			     (setq img-type (paste-image-type-from-file-name file-path))))))))
      (setq file-name file-path)))
  (let* ((cur-dir default-directory)
	 (dir-name (or (file-name-directory file-name) "./"))
	 (dir-name-noslash (directory-file-name dir-name))
	 (fname (concat (file-name-nondirectory file-name)
			(if img-type
			    (if (string-match (concat "\\." paste-image-type "$") file-name)
				""
			      (concat "." paste-image-type))
			  (progn (setq img-type "png") ".png")))))
    (if (file-exists-p dir-name-noslash)
	(unless (file-directory-p dir-name-noslash)
	  (error "%s is not a directory" dir-name))
      (make-directory dir-name-noslash))
    (if img-type
	(if (symbolp img-type)
	    (setq img-type (symbol-name img-type)))
      (error "Image type not set"))
    (paste-image-paste dir-name fname img-type)
    (insert paste-image-prefix (file-relative-name dir-name cur-dir) fname paste-image-postfix)
    (run-hooks 'paste-image-file-hook)))

(global-set-key (kbd "C-S-y") 'paste-image)

(provide 'paste-image)
;;; paste-image.el ends here
