import ftplib,easygui

# Replace with your FTP server address
ftp_server = "q1.cityhost.com.ua"

# Replace with your username and password (if required)
username = "chb9620ff2_dutko"
password = easygui.passwordbox("Enter password:")

# Specify the file to download on the server and where to save it locally
#ftp:///assets/2024/04/c1.txt
remote_file = "/assets/2024/04/c1.txt"
local_file = "c1_u.txt"

try:
  # Connect to the FTP server
  ftp = ftplib.FTP(ftp_server)
  
  # Login if username and password are required
  if username and password:
    ftp.login(username, password)
  
  # Open a local file for writing in binary mode
  with open(local_file, "wb") as local_handle:
    
    # Download the file from the server
    def download_callback(chunk):
      local_handle.write(chunk)
    
    ftp.retrbinary("RETR " + remote_file, download_callback)
  
  print(f"Downloaded file: {local_file}")
  
except ftplib.all_errors as e:
  print(f"FTP error: {e}")
finally:
  # Close the FTP connection
  if ftp:
    ftp.quit()


    import paramiko

def download_file_from_sftp(hostname, port, username, password, remote_filepath, local_filepath):
    try:
        # Create a transport object
        transport = paramiko.Transport((hostname, port))
        
        # Connect to the SSH server
        transport.connect(username=username, password=password)
        
        # Create an SFTP session
        sftp = paramiko.SFTPClient.from_transport(transport)
        
        # Download the file
        sftp.get(remote_filepath, local_filepath)
        
        # Close the SFTP session and transport
        sftp.close()
        transport.close()
        
        print("File downloaded successfully.")
    except Exception as e:
        print("Error:", e)

# Example usage
hostname = 'sftp.example.com'
port = 22  # Default port for SFTP is 22
username = 'your_username'
password = 'your_password'
remote_filepath = '/path/to/remote/file.txt'  # Specify the remote file path
local_filepath = 'local_file.txt'

download_file_from_sftp(hostname, port, username, password, remote_filepath, local_filepath)

