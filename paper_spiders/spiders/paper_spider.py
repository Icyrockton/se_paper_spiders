import scrapy
from scrapy.http import Response,Request
from ..utils.paperlist import paper_list


class PaperSpider(scrapy.Spider):
    name = 'paper_spider'

    def start_requests(self):
        for p in paper_list:

            yield Request(url=p['url'],callback=self.parse,cb_kwargs= p)


    def parse(self, response: Response, **kwargs):
        table = response.xpath('//*[@id="event-overview"]/table')
        papers = table.xpath('tr/td[2]/a[1]/text()').getall()
        for p in papers:
            yield {
                'title' : kwargs['title'],
                'paper' : p
            }
