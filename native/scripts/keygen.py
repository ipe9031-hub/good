"""Generate a local development strong-name key; not an Authenticode certificate."""
import pathlib,struct,sys
from cryptography.hazmat.primitives.asymmetric import rsa
path=pathlib.Path(sys.argv[1]);path.parent.mkdir(parents=True,exist_ok=True)
if path.exists(): print('Reusing local strong-name key');sys.exit(0)
p=rsa.generate_private_key(public_exponent=65537,key_size=2048).private_numbers()
def le(n,size):return n.to_bytes(size,'little')
data=struct.pack('<BBHI',7,2,0,0x2400)+struct.pack('<III',0x32415352,2048,p.public_numbers.e)
data+=le(p.public_numbers.n,256)+b''.join(le(x,128) for x in [p.p,p.q,p.dmp1,p.dmq1,p.iqmp])+le(p.d,256)
path.write_bytes(data);path.chmod(0o600);print('Created local development strong-name key')
