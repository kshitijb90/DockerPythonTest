FROM python
MAINTAINER ssriva41
RUN apt-get update && apt-get  install vim -y
RUN pip install xlrd
RUN pip install xlsxwriter
RUN pip install pyral
RUN mkdir /app
RUN chown 1001 /app
USER 1001
RUN cd /app; git clone https://github.optum.com/OptumRxAutomation/Pull_Rally_Reports.git
WORKDIR /app/Pull_Rally_Reports
#ENV FM_APP=FMtest.py
#ENV FM_ENV=development
EXPOSE 5000
#ENTRYPOINT ["python", "FMtest.py"]
ENTRYPOINT ["python"]
CMD ["Rally_Timebox_Reconcile_Dashboard.py", "Milestone", "safeApiKey"]
#CMD ["SprintVise_TTB.py", "safeApiKey"]
#CMD ["python","-m","test","run","--host=0.0.0.0"]
