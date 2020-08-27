FROM python
MAINTAINER ssriva41
RUN apt-get update && apt-get  install vim -y
RUN pip install xlrd
RUN mkdir /app
RUN chown 1001 /app
USER 1001
RUN cd /app; git clone https://github.optum.com/OptumRxAutomation/DockerPythonTest.git
WORKDIR /app/DockerPythonTest
ENV FM_APP=FMtest.py
ENV FM_ENV=development
EXPOSE 5000
ENTRYPOINT ["python", "FMtest.py"]
#CMD ["python","-m","test","run","--host=0.0.0.0"]
