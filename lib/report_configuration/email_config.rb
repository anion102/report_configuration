# encoding:utf-8
# author:anion
require 'mail'
module ReportConfiguration

  #send email
  module EmailConfig
    @@smtp=''
    def self.smtp_config(address,port,domain,name,pwd)
      @@smtp = { :address => address,
                 :port => port,
                 :domain => domain,
                 :user_name => name,
                 :password => pwd,
                 :enable_starttls_auto => true,
                 :openssl_verify_mode => 'none'
      }
    end

    def self.send_email(from,to,subject,body,file)
      Mail.defaults { delivery_method :smtp, @@smtp }
      mail = Mail.new do
        from from
        to to
        subject subject
        body body
        add_file File.expand_path(file)
      end
      mail.deliver!
    end

  end

end