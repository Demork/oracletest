# encoding : utf-8
import paramiko

class LinuxBase(object):
    def __init__(self,params):
        self.params = params
        self.hostname = self.params['hostname']
        self.port = self.params['port']
        self.username = self.params['username']
        self.password = self.params['password']

    @classmethod
    def convert_params(cls,params):
        params['port'] = int(params.get('port', 0))
        return params

    # ssh connection and sftp connection
    def connection(self):
        self.params = self.convert_params(self.params)
        ssh_client = paramiko.SSHClient()
        ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        try:
            ssh_client.connect(**self.params)
            t = paramiko.Transport((self.hostname,self.port))
            t.connect(username=self.username,password=self.password)
            sftp_client = paramiko.SFTPClient.from_transport(t)
            return ssh_client,sftp_client
        except Exception as e:
            print("linux connect error")
            return(None,None)
    # read all of file content
    def readfile(self,file,seek=0):
        _,sftp_client = self.connection()
        remote_file = sftp_client.open(file,'rb')
        remote_file.seek(seek)
        for line in remote_file.read().splitlines():
            if line != '':
                yield (line,0)
        yield (b'',remote_file.tell())
    # read last n of file content
    def readfile_n(self,file,num):
        cmd = 'tail -{} {}'.format(num,file)
        ssh_client,_ = self.connection()
        std_in, std_out, std_err = ssh_client.exec_command(cmd)
        for line in std_out.read().splitlines():
            if line != '':
                yield (line,0)

    def exec_command(self,command,ssh_client=None):
        try:
            if not ssh_client:
                ssh_client, _ = self.connection()
            std_in, std_out, std_err = ssh_client.exec_command(command)
            return std_out
        except Exception as e:
            print(e)

    def sftp_upload_file(self,local_file,remote_file):
        try:
            _, sftp_client = self.connection()
            sftp_client.put(local_file, remote_file)
        except Exception as e:
            print(e)

    def sftp_down_file(self,remote_file, local_file):
        try:
            _, sftp_client = self.connection()
            sftp_client.get(remote_file, local_file)
        except Exception as e:
            print(e)

