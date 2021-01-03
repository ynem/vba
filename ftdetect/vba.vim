" Notice that we didn't use an autocommand group like we usually would.
" Vim automatically wraps the contents of ftdetect/*.vim files in autocommand groups for you,
" so you don't need to worry about it.
autocmd BufNewFile,BufRead,BufEnter *.bas,*.cls,*.frm setfiletype vba

