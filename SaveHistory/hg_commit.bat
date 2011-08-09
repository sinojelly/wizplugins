@rem %1  hg path
@rem %2  work dir
@rem %3  comment

cd /d %2
%1\hg ci -A -m %3
%1\hg push --insecure


