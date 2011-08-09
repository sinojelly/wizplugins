@rem %1  git path
@rem %2  work dir
@rem %3  comment

cd /d %2
%1\git.exe add -A
%1\git.exe commit -m %3
%1\git.exe push
