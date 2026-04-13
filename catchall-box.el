;;; catchall-box.el --- UUID-based file management for Org mode -*- lexical-binding: t; -*-

;; Author: Fujisawa Electric Management Office
;; URL: https://github.com/zawatton21/catchall-box
;; Version: 0.1.0
;; Keywords: org, files, convenience
;; Package-Requires: ((emacs "27.1")) --- A utility for managing a catchall-box directory in Org mode  -*- lexical-binding: t; -*-

;;; Commentary:
;; 
;; This package allows you to manage a designated directory (the "catchall-box")
;; where files can be renamed to UUIDs automatically. The links to these files
;; are inserted into the current buffer in an Org-friendly format:
;;
;;   >> [[catchall-box:NEW_UUID.EXT][OLD_FILENAME.EXT]]
;;
;; It also provides a function to revert those filenames from UUID back to the
;; original names using the stored description in the link.
;;
;; Usage:
;;
;; 1. Adjust `catchall-box-directory` as desired (default is `~/Documents`).
;; 2. Call `catchall-box-update-link-abbrev` to update `org-link-abbrev-alist`.
;; 3. Use `M-x catchall-box-rename-to-uuid` to rename non-UUID
;;    files and insert links.
;; 4. Use `M-x catchall-box-revert-filenames` to revert the names back to
;;    the originals based on the link descriptions.
;;
;; For advanced usage, you can set `catchall-box-directory` conditionally depending
;; on your OS, then call `catchall-box-update-link-abbrev`.

;;; Code:

(require 'org-id)
(require 'rx)
(require 'subr-x)

(defgroup catchall-box nil
  "Settings for managing a personal catchall-box in Org mode."
  :group 'org
  :prefix "catchall-box-")

(defcustom catchall-box-directory (expand-file-name "~/Documents")
  "Directory used as the 'catchall-box' for files.

Defaults to `~/Documents`."
  :type 'directory
  :group 'catchall-box)

(defcustom catchall-box-onedrive-auto-detect nil
  "When non-nil, try to auto-detect the OneDrive root and set `catchall-box-directory` to it.

Specifically sets it to (OneDrive root)/(value of `catchall-box-onedrive-subdir`)."
  :type 'boolean
  :group 'catchall-box)

(defcustom catchall-box-onedrive-subdir "Catchall-Box"
  "Fixed subdirectory name under OneDrive to use as catchall box."
  :type 'string
  :group 'catchall-box)

;; ------------------------------------------------------------
;; Helpers
;; ------------------------------------------------------------

(defun catchall-box--file-exists-dir-p (path)
  (and (stringp path) (file-directory-p path)))

(defun catchall-box--trim-trailing-slash (path)
  (if (and (stringp path) (string-suffix-p "/" path))
      (substring path 0 -1)
    path))

(defun catchall-box--windows-path-to-wsl (win-path)
  "Convert like 'C:\\Users\\Name\\OneDrive' -> '/mnt/c/Users/Name/OneDrive'."
  (when (and win-path (string-match "\\`\\([A-Za-z]\\):\\(.*\\)\\'" win-path))
    (let* ((drive (downcase (match-string 1 win-path)))
           (rest  (match-string 2 win-path)))
      (concat "/mnt/" drive (subst-char-in-string ?\\ ?/ rest)))))

(defun catchall-box--path-to-file-uri (path)
  "Return proper file URI for PATH with trailing slash."
  (let* ((p (directory-file-name (expand-file-name path)))
         (win? (eq system-type 'windows-nt)))
    (cond
     ((and win? (string-match-p "\\`[A-Za-z]:" p))
      ;; Windows drive letter: ensure file:///C:/.../
      (concat "file:///" (replace-regexp-in-string "\\\\" "/" p) "/"))
     (t
      ;; POSIX
      (concat "file://" p "/")))))

(defun catchall-box--read-reg (key value)
  "Return REG_SZ of VALUE under registry KEY via reg.exe. Windows/WSL only."
  (let* ((exe (cond
               ((eq system-type 'windows-nt) "reg")
               ;; WSL 側から Windows の reg.exe を叩く
               ((and (eq system-type 'gnu/linux)
                     (or (bound-and-true-p *is-wsl*)
                         (getenv "WSL_DISTRO_NAME")))
                "/mnt/c/Windows/System32/reg.exe")
               (t nil))))
    (when exe
      (with-temp-buffer
        (let ((status (call-process exe nil (current-buffer) nil
                                    "query" key "/v" value)))
          (when (and (integerp status) (zerop status))
            (goto-char (point-min))
            (when (re-search-forward
                   (concat "\\b" (regexp-quote value) "\\b\\s-+REG_SZ\\s-+\\(.*\\)\\s-*$") nil t)
              (string-trim (match-string 1)))))))))

(defun catchall-box--onedrive-candidates-windows ()
  "Return list of possible OneDrive roots on Windows."
  (let* ((env (getenv "OneDrive"))
         (reg-main (catchall-box--read-reg "HKCU\\Software\\Microsoft\\OneDrive" "OneDrivePath"))
         (reg-pers (catchall-box--read-reg "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Personal" "UserFolder"))
         ;; Business1..Business9 ぐらいまでざっくり見る
         (biz (seq-some
               (lambda (n)
                 (catchall-box--read-reg
                  (format "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Business%d" n)
                  "UserFolder"))
               (number-sequence 1 9)))
         (userprofile (getenv "USERPROFILE"))
         (fallback (when userprofile (expand-file-name "OneDrive" userprofile))))
    (seq-filter #'catchall-box--file-exists-dir-p
                (delq nil (list env reg-main reg-pers biz fallback)))))

(defun catchall-box--onedrive-candidates-wsl ()
  "Return list of possible OneDrive roots when running in WSL."
  (let* ((reg-main (catchall-box--read-reg "HKCU\\Software\\Microsoft\\OneDrive" "OneDrivePath"))
         (reg-pers (catchall-box--read-reg "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Personal" "UserFolder"))
         (win-path (or reg-main reg-pers))
         (wsl-from-reg (when win-path (catchall-box--windows-path-to-wsl win-path)))
         (drives (mapcar (lambda (c) (format "/mnt/%c/Users" c))
                         (number-sequence ?c ?z)))
         (user-dirs (seq-mapcat
                     (lambda (ud)
                       (when (file-directory-p ud)
                         (directory-files ud t "^[^.]")))
                     drives))
         (probe (apply #'append
                       (mapcar (lambda (home)
                                 (list (expand-file-name "OneDrive" home)))
                               user-dirs))))
    (seq-filter #'catchall-box--file-exists-dir-p
                (delq nil (append (list wsl-from-reg) probe)))))

(defun catchall-box--onedrive-candidates-darwin ()
  "Return list of possible OneDrive roots on macOS."
  (let* ((cloud (expand-file-name "~/Library/CloudStorage"))
         (cloud-cands (when (file-directory-p cloud)
                        (directory-files cloud t "^OneDrive.*")))
         (legacy (directory-files (expand-file-name "~") t "^OneDrive.*")))
    (seq-filter #'catchall-box--file-exists-dir-p
                (append cloud-cands legacy))))

(defun catchall-box--detect-onedrive-root ()
  "Detect OneDrive root directory. Return a single existing path or nil."
  (cond
   ((eq system-type 'windows-nt)
    (car (catchall-box--onedrive-candidates-windows)))
   ((and (eq system-type 'gnu/linux)
         (or (bound-and-true-p *is-wsl*)
             (getenv "WSL_DISTRO_NAME")))
    (car (catchall-box--onedrive-candidates-wsl)))
   ((eq system-type 'darwin)
    (car (catchall-box--onedrive-candidates-darwin)))
   (t
    ;; Linux 通常環境などは固定クライアントを想定（必要なら拡張）
    (let ((home (expand-file-name "~/OneDrive")))
      (when (file-directory-p home) home)))))

;; ------------------------------------------------------------
;; Public API
;; ------------------------------------------------------------

;;;###autoload
(defun catchall-box-update-link-abbrev ()
  "Update `org-link-abbrev-alist` for the 'catchall-box' key from `catchall-box-directory`."
  (interactive)
  (setq org-link-abbrev-alist (assoc-delete-all "catchall-box" org-link-abbrev-alist))
  (let* ((dir (directory-file-name (expand-file-name catchall-box-directory)))
         (uri (catchall-box--path-to-file-uri dir)))
    (add-to-list 'org-link-abbrev-alist (cons "catchall-box" uri))))

;;;###autoload
(defun catchall-box-auto-configure ()
  "If auto-detect is enabled, set `catchall-box-directory` from OneDrive root and update link abbrev.
If detection fails, keep the existing `catchall-box-directory` as-is."
  (interactive)
  (when catchall-box-onedrive-auto-detect
    (let ((root (catchall-box--detect-onedrive-root)))
      (if (catchall-box--file-exists-dir-p root)
          (setq catchall-box-directory
                (expand-file-name
                 (concat (file-name-as-directory root)
                         (catchall-box--trim-trailing-slash catchall-box-onedrive-subdir))))
        (message "[catchall-box] OneDrive root not found. Using fallback: %s" catchall-box-directory))))
  (catchall-box-update-link-abbrev))

;;;###autoload
(defun catchall-box-insert-clipboard-link (description)
  "Insert a 'catchall-box' link using the current clipboard content.
Prompts for a DESCRIPTION to use as the link's text."
  (interactive "sEnter link description: ")
  (let ((filename (substring-no-properties (current-kill 0 t))))
    (insert (format "[[catchall-box:%s][%s]]" filename description))))

;;;###autoload
(defun catchall-box-rename-to-uuid ()
  "Recursively rename files in `catchall-box-directory` that are NOT UUID-named.
Renames them to uppercase UUIDs, then inserts links into the current buffer.

The inserted link format is:
  >> [[catchall-box:NEW_UUID.ext][OLD_FILENAME.ext]]"
  (interactive)
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist or is not set properly: %s"
           catchall-box-directory))
  (let* ((base-dir (expand-file-name catchall-box-directory))
         (files (directory-files-recursively base-dir ".*" t))
         (uuid-regexp
          "\\`[0-9A-Fa-f]\\{8\\}-[0-9A-Fa-f]\\{4\\}-[0-9A-Fa-f]\\{4\\}-[0-9A-Fa-f]\\{4\\}-[0-9A-Fa-f]\\{12\\}\\'"))
    (dolist (file files)
      (when (file-regular-p file)
        (let* ((fname (file-name-nondirectory file))
               (fdir  (file-name-directory file))
               (base  (file-name-sans-extension fname))
               (ext   (file-name-extension fname t)))
          (unless (string-match-p uuid-regexp (downcase base))
            (let* ((new-uuid (upcase (org-id-new)))
                   (old-fname fname)
                   (new-full (concat fdir new-uuid ext)))
              (rename-file file new-full t)
              (insert (format ">> [[catchall-box:%s%s][%s]]\n"
                              new-uuid ext old-fname)))))))))

;;;###autoload
(defun catchall-box-revert-filenames (start end)
  "Revert UUID-based filenames to their original names using links in the selected region.

Looks for lines of the form:
  >> [[catchall-box:NEWFILENAME.ext][OLDFILENAME.ext]]
and renames the actual file from NEWFILENAME.ext to OLDFILENAME.ext
within `catchall-box-directory`."
  (interactive "r")
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist or is not set properly: %s"
           catchall-box-directory))
  (let* ((base-dir (expand-file-name catchall-box-directory))
         (text (buffer-substring-no-properties start end)))
    (with-temp-buffer
      (insert text)
      (goto-char (point-min))
      (while (re-search-forward
              (rx bol ">> " "[[catchall-box:" (group (+ (not (any "]"))))
                  "][" (group (+ (not (any "]")))) "]]")
              nil t)
        (let* ((new-filename (match-string 1))
               (old-filename (match-string 2))
               (old-path (expand-file-name new-filename base-dir))
               (new-path (expand-file-name old-filename base-dir)))
          (if (file-exists-p old-path)
              (progn
                (rename-file old-path new-path t)
                (message "Renamed: %s -> %s" old-path new-path))
            (message "File not found (skip): %s" old-path)))))))

;; ------------------------------------------------------------
;; Store file into Catchall-Box
;; ------------------------------------------------------------

;;;###autoload
(defun catchall-box-store-file (file &optional description)
  "Copy FILE into `catchall-box-directory` with a UUID-based name.
Return the new filename (UUID.ext) as a string.
If DESCRIPTION is nil, use the original filename."
  (interactive
   (list (read-file-name "File to store in Catchall-Box: " nil nil t)))
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist: %s" catchall-box-directory))
  (let* ((orig-name (file-name-nondirectory file))
         (ext (file-name-extension orig-name t))  ; includes "."
         (uuid (upcase (org-id-new)))
         (new-name (concat uuid ext))
         (dest (expand-file-name new-name catchall-box-directory))
         (desc (or description orig-name)))
    (copy-file file dest nil t t)
    (message "Stored: %s -> %s" orig-name new-name)
    ;; When called interactively, also insert a link at point
    (when (called-interactively-p 'any)
      (insert (format ">> [[catchall-box:%s][%s]]\n" new-name desc)))
    new-name))

;;;###autoload
(defun catchall-box-make-link (file-in-box &optional description)
  "Insert a catchall-box link for FILE-IN-BOX (a filename already in the box).
Prompts with completion from existing files in `catchall-box-directory`."
  (interactive
   (let* ((files (directory-files catchall-box-directory nil "^[^.]"))
          (chosen (completing-read "File in Catchall-Box: " files nil t)))
     (list chosen (read-string "Description: " chosen))))
  (let ((desc (or description file-in-box)))
    (insert (format ">> [[catchall-box:%s][%s]]" file-in-box desc))))

;; ------------------------------------------------------------
;; Store URL article into Catchall-Box
;; ------------------------------------------------------------

(defun catchall-box--fetch-url-title (url)
  "Fetch the <title> of URL. Return title string or nil."
  (condition-case nil
      (with-temp-buffer
        (url-insert-file-contents url)
        (goto-char (point-min))
        (when (re-search-forward "<title[^>]*>\\([^<]+\\)</title>" nil t)
          (string-trim (match-string 1))))
    (error nil)))

;;;###autoload
(defun catchall-box-store-url (url)
  "Download the web page at URL, store it in Catchall-Box as UUID.html.
Insert a catchall-box link and a source URL line at point.

Result in buffer:
  >> [[catchall-box:UUID.html][Page Title]]
  src: URL"
  (interactive "sURL: ")
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist: %s" catchall-box-directory))
  (let* ((title (or (catchall-box--fetch-url-title url)
                     (read-string "Title (could not auto-detect): ")))
         (uuid (upcase (org-id-new)))
         (new-name (concat uuid ".html"))
         (dest (expand-file-name new-name catchall-box-directory)))
    ;; Download full HTML
    (url-copy-file url dest)
    (insert (format ">> [[catchall-box:%s][%s]]\nsrc: %s\n" new-name title url))
    (message "Stored URL: %s -> %s" url new-name)))

;; ------------------------------------------------------------
;; Delete file from catchall-box link at point
;; ------------------------------------------------------------

(defun catchall-box--link-at-point ()
  "Return (FILENAME . DESCRIPTION) of catchall-box link at point, or nil.
Handles both unexpanded (catchall-box:UUID.ext) and expanded
\(file:///path/to/Catchall-Box/UUID.ext) link forms."
  (let ((ctx (org-element-context)))
    (when (eq (org-element-type ctx) 'link)
      (let* ((raw-path (org-element-property :raw-link ctx))
             (path (org-element-property :path ctx))
             (desc (when-let ((cb (org-element-property :contents-begin ctx))
                              (ce (org-element-property :contents-end ctx)))
                     (buffer-substring-no-properties cb ce)))
             (filename
              (cond
               ;; Case 1: unexpanded abbreviation  catchall-box:UUID.ext
               ((string-prefix-p "catchall-box:" raw-path)
                (substring raw-path (length "catchall-box:")))
               ;; Case 2: org-element expanded the abbreviation to full path
               ((and (stringp path)
                     (let* ((box (downcase (file-name-as-directory
                                           (expand-file-name catchall-box-directory))))
                            (p   (downcase (expand-file-name path))))
                       (string-prefix-p box p)))
                (file-name-nondirectory path))
               (t nil))))
        (when filename
          (cons filename desc))))))

;;;###autoload
(defun catchall-box-delete-file-at-point ()
  "Delete the Catchall-Box file referenced by the org link at point.
Also removes the link line from the buffer."
  (interactive)
  (let ((link-info (catchall-box--link-at-point)))
    (unless link-info
      (user-error "No catchall-box link at point"))
    (let* ((filename (car link-info))
           (desc (or (cdr link-info) filename))
           (full-path (expand-file-name filename catchall-box-directory)))
      (unless (file-exists-p full-path)
        (user-error "File does not exist: %s" full-path))
      (when (y-or-n-p (format "Delete '%s' (%s)? " desc filename))
        (delete-file full-path)
        ;; Remove the whole line containing the link
        (save-excursion
          (beginning-of-line)
          (let ((beg (point)))
            (forward-line 1)
            (delete-region beg (point))))
        (message "Deleted: %s" filename)))))

;;;###autoload
(defun catchall-box-delete-files-in-region (beg end)
  "Delete all Catchall-Box files referenced by links in the region BEG..END.
Also removes the corresponding link lines from the buffer.
Prompts with a summary before proceeding."
  (interactive "r")
  (unless (use-region-p)
    (user-error "No region selected"))
  (save-excursion
    (save-restriction
      (narrow-to-region beg end)
      ;; Collect all catchall-box links in the region
      (goto-char (point-min))
      (let ((links nil))
        (while (re-search-forward
                "\\[\\[catchall-box:\\([^]]+\\)\\]\\[\\([^]]+\\)\\]\\]" nil t)
          (let ((filename (match-string-no-properties 1))
                (desc (match-string-no-properties 2))
                (line-beg (save-excursion (beginning-of-line) (point)))
                (line-end (save-excursion (forward-line 1) (point))))
            (push (list filename desc line-beg line-end) links)))
        (when (null links)
          (user-error "No catchall-box links found in region"))
        (setq links (nreverse links))
        ;; Show summary and confirm
        (let* ((existing
                (seq-filter (lambda (l)
                              (file-exists-p
                               (expand-file-name (nth 0 l) catchall-box-directory)))
                            links))
               (missing (- (length links) (length existing))))
          (when (null existing)
            (user-error "None of the %d linked files exist in Catchall-Box"
                        (length links)))
          (when (y-or-n-p
                 (format "Delete %d file(s) from Catchall-Box?%s "
                         (length existing)
                         (if (> missing 0)
                             (format " (%d already missing)" missing)
                           "")))
            ;; Delete files and remove lines (process in reverse to keep positions valid)
            (let ((deleted 0))
              (dolist (entry (reverse existing))
                (let* ((filename (nth 0 entry))
                       (line-beg (nth 2 entry))
                       (line-end (nth 3 entry))
                       (full-path (expand-file-name filename catchall-box-directory)))
                  (ignore-errors (delete-file full-path t))
                  (delete-region line-beg line-end)
                  (cl-incf deleted)))
              (message "Deleted %d file(s) from Catchall-Box" deleted))))))))

;; ------------------------------------------------------------
;; Orphan file detection
;; ------------------------------------------------------------

;;;###autoload
(defun catchall-box-orphan-files ()
  "List files in Catchall-Box that are not referenced by any org file.
Searches all .org files in `org-mem-watch-dirs' (or `org-directory')
for catchall-box links and compares against actual files."
  (interactive)
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist: %s" catchall-box-directory))
  (let* ((box-files (directory-files catchall-box-directory nil "^[^.]"))
         ;; Collect all catchall-box references from org files
         (org-dirs (if (bound-and-true-p org-mem-watch-dirs)
                       org-mem-watch-dirs
                     (list (or (bound-and-true-p org-directory) "~/org/"))))
         (org-files (seq-mapcat
                     (lambda (dir)
                       (when (file-directory-p dir)
                         (directory-files-recursively dir "\\.org\\'")))
                     org-dirs))
         (referenced (make-hash-table :test 'equal)))
    ;; Scan all org files for catchall-box: references
    (dolist (f org-files)
      (with-temp-buffer
        (insert-file-contents f)
        (goto-char (point-min))
        (while (re-search-forward "catchall-box:\\([^][\n]+\\)" nil t)
          (puthash (match-string 1) t referenced))))
    ;; Find orphans
    (let ((orphans (seq-remove (lambda (name) (gethash name referenced)) box-files)))
      (if (null orphans)
          (message "No orphan files found in Catchall-Box.")
        (with-current-buffer (get-buffer-create "*Catchall-Box Orphans*")
          (erase-buffer)
          (insert (format "Orphan files in Catchall-Box (%d found):\n\n" (length orphans)))
          (dolist (name orphans)
            (insert (format "  %s\n" name)))
          (insert (format "\nCatchall-Box: %s\n" catchall-box-directory))
          (goto-char (point-min))
          (display-buffer (current-buffer)))
        (message "%d orphan file(s) found." (length orphans))))))

;; ------------------------------------------------------------
;; Copy single file at point to shared folder
;; ------------------------------------------------------------

(defcustom catchall-box-export-directory
  (expand-file-name "~/Desktop")
  "Default directory for `catchall-box-copy-file-at-point' and
`catchall-box-export-subtree'.
Typically overridden by `catchall-box-auto-configure' in init.el
to use the OneDrive Desktop folder."
  :type 'directory
  :group 'catchall-box)

(defun catchall-box--sanitize-filename (name)
  "Remove characters that are invalid in file/directory names."
  (replace-regexp-in-string "[<>:\"/\\\\|?*\x00-\x1f]" "_" name))

;;;###autoload
(defun catchall-box-copy-file-at-point (&optional dest-dir)
  "Copy the Catchall-Box file at point to DEST-DIR with its human-readable name.

Uses the link description as the destination filename.
DEST-DIR defaults to `catchall-box-export-directory' (typically ~/Desktop).

Example: with cursor on
  >> [[catchall-box:01992E61-09CF-78D9-8D8C-20055E4D7F00.pdf][重要なお知らせ.pdf]]

copies the UUID file to ~/Desktop/重要なお知らせ.pdf"
  (interactive
   (list (read-directory-name "Copy to: " catchall-box-export-directory)))
  (let ((link-info (catchall-box--link-at-point)))
    (unless link-info
      (user-error "No catchall-box link at point"))
    (let* ((uuid-name (car link-info))
           (desc (or (cdr link-info) uuid-name))
           (dest (or dest-dir catchall-box-export-directory))
           ;; Ensure description keeps the correct extension
           (desc-ext (file-name-extension desc))
           (uuid-ext (file-name-extension uuid-name))
           (final-name
            (if desc-ext
                (catchall-box--sanitize-filename desc)
              (catchall-box--sanitize-filename
               (if uuid-ext (concat desc "." uuid-ext) desc))))
           (src (expand-file-name uuid-name catchall-box-directory))
           (dst (expand-file-name final-name dest)))
      (unless (file-exists-p src)
        (user-error "File does not exist: %s" src))
      (make-directory dest t)
      (copy-file src dst t)
      (message "Copied: %s -> %s" uuid-name dst)
      dst)))

;; ------------------------------------------------------------
;; Export subtree as folder with renamed files
;; ------------------------------------------------------------

(defun catchall-box--collect-subtree-links ()
  "Parse the subtree at point and return a list of (REL-DIR FILENAME DESCRIPTION).
REL-DIR is the relative folder path derived from sub-heading hierarchy.
FILENAME is the catchall-box filename (UUID.ext).
DESCRIPTION is the link description (human-readable name)."
  (save-excursion
    (org-back-to-heading t)
    (let* ((top-level (org-current-level))
           (end (save-excursion (org-end-of-subtree t t) (point)))
           (results nil)
           ;; Stack of (level . heading) to track folder path
           (heading-stack nil))
      ;; The top heading itself becomes the root folder name (handled by caller)
      (forward-line 1)
      (while (< (point) end)
        (cond
         ;; Sub-heading: update the heading stack
         ((looking-at org-heading-regexp)
          (let ((level (org-current-level))
                (title (org-get-heading t t t t)))
            ;; Remove headings that are at same or deeper level
            (setq heading-stack
                  (seq-take-while (lambda (pair) (< (car pair) level))
                                  heading-stack))
            (push (cons level (catchall-box--sanitize-filename title))
                  heading-stack)))
         ;; Line with catchall-box link
         ((looking-at ".*\\[\\[catchall-box:\\([^]]+\\)\\]\\[\\([^]]+\\)\\]\\]")
          (let* ((filename (match-string-no-properties 1))
                 (desc (match-string-no-properties 2))
                 ;; Build relative path from heading stack (reversed = top-to-bottom)
                 (path-parts (mapcar #'cdr (reverse heading-stack)))
                 (rel-dir (if path-parts
                              (mapconcat #'identity path-parts "/")
                            "")))
            (push (list rel-dir filename desc) results))))
        (forward-line 1))
      (nreverse results))))

;;;###autoload
(defun catchall-box-export-subtree (&optional dest-dir)
  "Export the subtree at point as a folder with human-readable filenames.

Creates a folder named after the current heading under DEST-DIR
\(default: `catchall-box-export-directory', i.e. Desktop).
Sub-headings become subdirectories.  Each catchall-box link's file
is copied and renamed from its UUID name to the link description.

Example: with cursor on \"* 過電流継電器 OCR\" containing
  ** 取扱説明書
     >> [[catchall-box:UUID.pdf][三菱MOC 取説.pdf]]

creates:
  ~/Desktop/過電流継電器 OCR/取扱説明書/三菱MOC 取説.pdf"
  (interactive
   (list (read-directory-name "Export to: " catchall-box-export-directory)))
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist: %s" catchall-box-directory))
  (let* ((dest (or dest-dir catchall-box-export-directory))
         ;; Get the top heading title for root folder name
         (root-name (save-excursion
                      (org-back-to-heading t)
                      (catchall-box--sanitize-filename
                       (org-get-heading t t t t))))
         (root-dir (expand-file-name root-name dest))
         (links (catchall-box--collect-subtree-links))
         (copied 0)
         (skipped 0))
    (when (null links)
      (user-error "No catchall-box links found in this subtree"))
    ;; Create root folder
    (make-directory root-dir t)
    ;; Process each link
    (dolist (entry links)
      (let* ((rel-dir (nth 0 entry))
             (uuid-name (nth 1 entry))
             (desc (nth 2 entry))
             ;; Ensure description keeps the correct extension
             (desc-ext (file-name-extension desc))
             (uuid-ext (file-name-extension uuid-name))
             (final-name
              (if desc-ext
                  ;; Description already has extension
                  (catchall-box--sanitize-filename desc)
                ;; Append UUID file's extension to description
                (catchall-box--sanitize-filename
                 (if uuid-ext
                     (concat desc "." uuid-ext)
                   desc))))
             (target-dir (if (string-empty-p rel-dir)
                             root-dir
                           (expand-file-name rel-dir root-dir)))
             (src (expand-file-name uuid-name catchall-box-directory))
             (dst (expand-file-name final-name target-dir)))
        (make-directory target-dir t)
        (if (file-exists-p src)
            (progn
              (copy-file src dst t)
              (cl-incf copied))
          (message "Skip (not found): %s" uuid-name)
          (cl-incf skipped))))
    (message "Exported %d file(s) to %s (skipped: %d)" copied root-dir skipped)
    root-dir))

;;;###autoload
(defun catchall-box--uuid-name-p (name)
  "Return non-nil if NAME (sans extension) looks like a UUID."
  (let ((base (file-name-sans-extension name)))
    (string-match-p
     "\\`[0-9A-Fa-f]\\{8\\}-[0-9A-Fa-f]\\{4\\}-[0-9A-Fa-f]\\{4\\}-[0-9A-Fa-f]\\{4\\}-[0-9A-Fa-f]\\{12\\}\\'"
     base)))

(defun catchall-box--import-dir-recursive (dir level)
  "Recursively build org text for DIR at heading LEVEL.
Each file is MOVED into `catchall-box-directory' root with a UUID name.
Return (TEXT . FILE-COUNT)."
  (let ((entries (directory-files dir t "^[^.]"))
        (text "")
        (count 0))
    (let ((dirs (sort (seq-filter #'file-directory-p entries)
                      #'string<))
          (files (sort (seq-filter #'file-regular-p entries)
                       #'string<)))
      ;; Files -> move to catchall-box root with UUID, insert links
      (dolist (file files)
        (let* ((orig-name (file-name-nondirectory file))
               (ext (file-name-extension orig-name t))
               (uuid (upcase (org-id-new)))
               (new-name (concat uuid ext))
               (dest (expand-file-name new-name catchall-box-directory)))
          (rename-file file dest nil)
          (setq text (concat text
                             (format ">> [[catchall-box:%s][%s]]\n"
                                     new-name orig-name)))
          (cl-incf count)))
      ;; Subdirectories -> sub-headings (recursive)
      (dolist (subdir dirs)
        (let* ((subdir-name (file-name-nondirectory subdir))
               (result (catchall-box--import-dir-recursive
                        subdir (1+ level))))
          (setq text (concat text
                             (make-string level ?*) " " subdir-name "\n"
                             (car result)))
          (cl-incf count (cdr result)))))
    (cons text count)))

;;;###autoload
(defun catchall-box-import-directory ()
  "Import all non-UUID-named folders in `catchall-box-directory' at point.
Each folder becomes a heading; its structure is preserved as sub-headings.
Files are moved to catchall-box root with UUID names and linked.
After import, the source folders are deleted.

The heading level is determined by context:
- Inside a heading: one level deeper (child)
- Top level: level 1

This is the inverse of `catchall-box-export-subtree'.

Workflow:
  1. Drop folders into Catchall-Box (OneDrive/Catchall-Box/)
  2. Place cursor in an org buffer
  3. C-c b i  — all non-UUID folders are detected and imported

Example: Catchall-Box/ contains
  過電流継電器 OCR/
    取扱説明書/
      三菱MOC 取説.pdf
  地絡方向継電器 DGR/
    試験成績書.pdf

produces at point:
  ** 過電流継電器 OCR
  *** 取扱説明書
  >> [[catchall-box:UUID1.pdf][三菱MOC 取説.pdf]]
  ** 地絡方向継電器 DGR
  >> [[catchall-box:UUID2.pdf][試験成績書.pdf]]"
  (interactive)
  (unless (file-directory-p catchall-box-directory)
    (error "Catchall-box directory does not exist: %s"
           catchall-box-directory))
  (let* ((all-entries (directory-files catchall-box-directory t "^[^.]"))
         (folders (seq-filter
                   (lambda (e)
                     (and (file-directory-p e)
                          (not (catchall-box--uuid-name-p
                                (file-name-nondirectory e)))))
                   all-entries)))
    (when (null folders)
      (user-error "No non-UUID folders found in %s" catchall-box-directory))
    (let* ((base-level (or (org-current-level) 0))
          (insert-level (1+ base-level))
          (total 0)
          (failed-dirs nil))
      ;; Ensure we start on a fresh line
      (unless (bolp) (insert "\n"))
      ;; Process each folder
      (dolist (folder (sort folders #'string<))
        (let* ((folder-name (file-name-nondirectory folder))
               (result (catchall-box--import-dir-recursive
                        folder (1+ insert-level)))
               (body (car result))
               (count (cdr result)))
          ;; Insert heading + body
          (insert (make-string insert-level ?*) " " folder-name "\n"
                  body)
          (cl-incf total count)
          ;; Remove the source folder
          ;; On Windows/OneDrive, folders get system/hidden/readonly attrs
          ;; that block deletion.  Strip attrs first, then force-remove.
          (condition-case err
              (if (eq system-type 'windows-nt)
                  (let ((win-path (replace-regexp-in-string "/" "\\\\" folder)))
                    ;; 1) Strip readonly/system/hidden attrs recursively
                    (call-process "cmd.exe" nil nil nil
                                  "/c" "attrib" "-r" "-s" "-h"
                                  (concat win-path "\\\\*") "/s" "/d")
                    (call-process "cmd.exe" nil nil nil
                                  "/c" "attrib" "-r" "-s" "-h" win-path)
                    ;; 2) Force-remove
                    (let ((ret (call-process "cmd.exe" nil nil nil
                                             "/c" "rd" "/s" "/q" win-path)))
                      (when (and (numberp ret) (not (zerop ret)))
                        ;; 3) Fallback: PowerShell Remove-Item
                        (let ((ps-ret (call-process
                                       "powershell.exe" nil nil nil
                                       "-NoProfile" "-Command"
                                       (format "Remove-Item -LiteralPath '%s' -Recurse -Force -ErrorAction Stop"
                                               (replace-regexp-in-string "\\\\\\\\" "\\\\" win-path)))))
                          (when (and (numberp ps-ret) (not (zerop ps-ret)))
                            (push (format "%s: delete failed (rd=%d ps=%d)"
                                          (file-name-nondirectory folder) ret ps-ret)
                                  failed-dirs))))))
                ;; Non-Windows: standard recursive delete
                (delete-directory folder t))
            (error
             (push (format "%s: %s" (file-name-nondirectory folder)
                           (error-message-string err))
                   failed-dirs)))))
      (if failed-dirs
          (message "Imported %d file(s) from %d folder(s). Could not remove: %s"
                   total (length folders)
                   (mapconcat #'identity failed-dirs "; "))
        (message "Imported %d file(s) from %d folder(s) (all source folders removed)"
                 total (length folders))))))

;; ------------------------------------------------------------
;; Backward compatibility aliases
;; ------------------------------------------------------------

(defalias 'insert-clipboard-as-catchall-link #'catchall-box-insert-clipboard-link)
(defalias 'rename-files-to-uuid-and-insert-catchall-links #'catchall-box-rename-to-uuid)
(defalias 'revert-filenames-from-catchall-links #'catchall-box-revert-filenames)

(provide 'catchall-box)
;;; catchall-box.el ends here
