@rem %1  local path
@rem %2  comment

cd /d %1
hg ci -A -m %2
hg push --insecure


